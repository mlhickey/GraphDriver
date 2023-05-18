//#define signinactivity
using Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace Services
{
    public class GuestUserServices
    {
        private readonly GraphServiceClient _client;
        private readonly List<string> _exemptionList;

        private readonly int _staleRange;
        private readonly int _staleInviteRange;
        private readonly int _removalRange;
        private readonly int _maxAttempts;
        private readonly string _guestProperties;

        private Func<IDictionary<string, object>, string, string> GetDictionaryValue = (x, y) => x.ContainsKey(y) ? x[y].ToString() : "0";

        public GuestUserServices(/*IGraphServiceClient graphServiceClient, IConfiguration configuration*/)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("local.settings.json")
                .Build();
            _client = new GraphServiceClientService(configuration).client;

            _guestProperties = configuration.GetValue<string>("GuestUserProps", GuestUserConfiguration.StaleGuestProperties);
            _maxAttempts = configuration.GetValue<int>("MaxAttempts", 8);
            //
            // Set removal range values to user staleness evaluation
            //
            _staleRange = SetValidRange(configuration.GetValue<int>("StaleRange", GuestUserConfiguration.StaleRange));
            _staleInviteRange = SetValidRange(configuration.GetValue<int>("StaleInvitationRange", GuestUserConfiguration.StaleRange));
            _removalRange = SetValidRange(configuration.GetValue<int>("RemovalRange", GuestUserConfiguration.RemovalRange));
            //
            // Check for exemption group, build list of value is non-null
            //
            var eGuid = configuration.GetValue<string>("ExemptionGroupGUID", GuestUserConfiguration.ExemptionGroupGUID);
#if by_guid
            _exemptionList = eGuid.Split(';').ToList();
#else
            if (!string.IsNullOrEmpty(eGuid))
            {
                _exemptionList = GetExemptionGroupMembership(eGuid).Result;
            }
            else
            {
                _exemptionList = new List<string>();
            }
#endif
            Console.WriteLine($"Total of {_exemptionList.Count} exempt users");
        }

        /// <summary>
        /// SetValidRange - ensures that associated range isn't below minimum threshold
        /// </summary>
        /// <param name="value"></param>
        /// <returns>
        /// Valid int value for use with AddDays method
        /// </returns>
        private int SetValidRange(int value)
        {
            if (value < GuestUserConfiguration.StaleRange)
            {
                Console.WriteLine($"Stale range {value} is below minimum of {GuestUserConfiguration.StaleRange}");
                value = GuestUserConfiguration.StaleRange;
            }
            return value * -1;
        }

        #region ExemptUsers

        /// <summary>
        /// IsMemberOfExceptionGroups - check to see if specified id is a member of an exception group
        /// </summary>
        /// <param name="id"></param>
        /// <returns>
        /// bool
        /// </returns>
        public bool IsMemberOfExceptionGroups(string id)
        {
            if (_exemptionList.Count == 0)
            {
                Console.WriteLine("No exemption group specified, all users will be in scope");
                return false;
            }
#if by_guid

            try
            {
                var res = await _client.Users[id].CheckMemberGroups(_exemptionList).Request().PostAsync();
                return res.Count > 0;
            }
            catch (ServiceException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                return true;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"IsMemberOfExceptionGroups: {ex.Message}");
                return false;
            }
#else

            return _exemptionList.Contains(id);
#endif
        }

        /// <summary>
        /// GetExemptionGroupMembership - Creates user ID list basded on transitive membership
        /// </summary>
        /// <returns>List<string></returns>
        private async Task<List<string>> GetExemptionGroupMembership(string groupId)
        {
            try
            {
                var exemptionMembers = await _client.Groups[groupId]
                    .TransitiveMembers
                    .Request()
                    .Select("id")
                    .WithMaxRetry(_maxAttempts)
                    .Top(999)
                    .GetAsync();

                var exemptionList = await ProcessRequestList(exemptionMembers);
                return exemptionList;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"GetExemptionGroupMembership retrival failed: {ex.InnerException.Message}");
                return new List<string>();
            }
        }

        #endregion ExemptUsers

        #region InactiveUsers

        /// <summary>
        /// GetInactiveGuests - Retrieves list of active guest accounts which are:
        ///     outside staleness threshold
        ///     not members of the exemption group
        /// </summary>
        /// <returns></returns>
        public async Task<List<User>> GetInactiveGuests()
        {
            return await GetGuestAccounts(true, _staleRange);
        }

        /// <summary>
        /// GetDisableInactiveGuests - Retrieves list of disabled guest accounts which are:
        ///     outside removal threshold
        ///     not members of the exemption group
        /// </summary>
        /// <returns></returns>
        public async Task<List<User>> GetDisableInactiveGuests()
        {
            return await GetGuestAccounts(false, _removalRange);
        }

        private async Task<List<User>> GetGuestAccounts(bool enabled, int staleRange, [CallerMemberName] string callerName = "")
        {
            var staleDate = DateTime.Now.AddDays(staleRange);
            var result = new List<User>();
            var isEnabled = enabled.ToString().ToLower();
            int oos = 0;

            var queryBuilder = new List<string> {
                "externalUserState eq 'Accepted'",
                "userType eq 'Guest'",
                $"accountEnabled eq {isEnabled}"
                };
            var queryFilter = string.Join(" and ", queryBuilder);

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("$count", "true")
            };
            Console.WriteLine($"{callerName}::GetGuestAccounts: Processing started at {DateTime.Now}");

#if signinactivity
            queryFilter = $"signInActivity/lastSignInDateTime le {staleDate.ToString("yyyy-MM-ddTHH:mm:ssZ")}";
#endif
            try
            {
                var request = await _client.Users
                   .Request(queryOptions)
                   .Filter(queryFilter)
                   .WithMaxRetry(_maxAttempts)
                   .Top(999)
                   .GetAsync();
#if signinactivity

                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    if ((u.UserType == "Guest" && u.ExternalUserState == "Accepted")
                    && u.AccountEnabled == enabled
                    ) result.Add(u);
                    return true;
                });
                await pageIterator.IterateAsync();
#else
                Console.WriteLine($"{callerName}::GetGuestAccounts: Total of {GetDictionaryValue(request.AdditionalData, "@odata.count")} {enabled} objects");
#if local
                result = await ProcessBoundRequestList(request, staleDate);
#else
                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    if (!IsMemberOfExceptionGroups(u.Id) && IsInactivePastThreshold(u, staleDate))
                        result.Add(u);
                    else
                        oos++;
                    return true;
                });
                await pageIterator.IterateAsync();
#endif
                Console.WriteLine($"{callerName}::GetGuestAccounts:Excluded {oos} users");
#endif
            }
            catch (ServiceException ex)
            {
                var failureType = result.Count() > 0 ? "partially" : "completely";
                Console.WriteLine($"{callerName}::GetGuestAccounts: retrieval {failureType} failed: {ex.InnerException.Message}");
            }

            return result;
        }

        #endregion InactiveUsers

        #region UnaccptedInvitations

        public async Task<List<User>> GetUnacceptedInvitees()
        {
            var result = new List<User>();
            var staleDate = DateTime.Now.AddDays(_staleInviteRange).ToString("yyyy-MM-ddTHH:mm:ssZ");

            var queryBuilder = new List<string> {
                "externalUserState eq 'PendingAcceptance'",
                "userType eq 'Guest'",
                $"createdDateTime le {staleDate}"
                };

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("$count", "true")
            };

            try
            {
                var request = await _client.Users
                   .Request(queryOptions)
                   .Filter(string.Join(" and ", queryBuilder))
                   .Select(_guestProperties)
                   .Top(999)
                   .GetAsync();
                Console.WriteLine($"{GetType().Name}: Total of {GetDictionaryValue(request.AdditionalData, "@odata.count")} objects");
#if local
                result = await ProcessRequestList(request);
#else
                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    result.Add(u); return true;
                });
                await pageIterator.IterateAsync();
#endif
            }
            catch (ServiceException ex)
            {
                var failureType = result.Count() > 0 ? "partially" : "completely";
                Console.WriteLine($"GetUnacceptedInvitees retrieval {failureType} failed: {ex.InnerException.Message}");
            }
            return result;
        }

        #endregion UnaccptedInvitations

        #region StalenessValidation

        /// <summary>
        /// GetLastSignIn evalautes datetime values to determine actual last signin period.
        /// Uses:
        ///     LastSignInDateTime
        ///     LastNonInteractiveSignInDateTime
        ///     CreatedDateTime
        ///     SignInSessionsValidFromDateTime
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public DateTimeOffset? GetLastSignIn(User user)
        {
            if (user?.SignInActivity != null)
                return (DateTimeOffset)user.SignInActivity?.LastSignInDateTime;

            return user?.SignInSessionsValidFromDateTime ?? user?.CreatedDateTime;
        }

        /// <summary>
        /// IsInactivePastThreshold - Checks if last login of user is outside of staleness threshold
        /// </summary>
        /// <param name="user"></param>
        /// <param name="staleDate"></param>
        /// <returns>bool</returns>
        private bool IsInactivePastThreshold(User user, DateTime staleDate)
        {
            var lastSignInDate = GetLastSignIn(user);
            if (lastSignInDate == null) return false;
            return DateTime.Compare(lastSignInDate.Value.DateTime, staleDate) < 0;
        }

        #endregion StalenessValidation

        #region Iterators

        /// <summary>
        /// ProcessRequestList performs NextPage processing of associated request
        /// </summary>
        /// <param name="request"></param>
        /// <returns>
        /// List of user objects returned fron NextPage processing
        /// </returns>
        private async Task<List<string>> ProcessRequestList(IGroupTransitiveMembersCollectionWithReferencesPage request)
        {
            var result = new List<string>();
            result.AddRange(request.CurrentPage.Select(i => i.Id));

            while (request.NextPageRequest != null)
            {
                request = await request.NextPageRequest
                    .WithMaxRetry(_maxAttempts)
                    .GetAsync();

                result.AddRange(request.CurrentPage.Select(i => i.Id));
            };
            return result.Distinct().ToList();
        }

        /// <summary>
        /// ProcessRequestList performs NextPage processing of associated request
        /// </summary>
        /// <param name="request"></param>
        /// <returns>
        /// List of user objects returned fron NextPage processing
        /// </returns>
        private async Task<List<User>> ProcessRequestList(IGraphServiceUsersCollectionPage request)
        {
            var result = new List<User>();
            result.AddRange(request.CurrentPage);

            while (request.NextPageRequest != null)
            {
                request = await request.NextPageRequest
                    .WithMaxRetry(_maxAttempts)
                    .GetAsync();

                result.AddRange(request.CurrentPage);
            };
            return result;
        }

        /// <summary>
        /// ProcessBoundRequestList performs NextPage processing of associated request with additional validation:
        ///     - Not member of exemption group
        ///     - Inactivity based on specified threshold date
        /// </summary>
        /// <param name="request"></param>
        /// <returns>
        /// List of valid user objects returned fron NextPage processing
        /// </returns>
        private async Task<List<User>> ProcessBoundRequestList(IGraphServiceUsersCollectionPage request, DateTime staleDate)
        {
            var result = new List<User>();

            result.AddRange(request.CurrentPage
                .Where(u => !IsMemberOfExceptionGroups(u.Id) && IsInactivePastThreshold(u, staleDate)));

            while (request.NextPageRequest != null)
            {
                try
                {
                    request = await request.NextPageRequest
                        .WithMaxRetry(_maxAttempts)
                        .GetAsync();

                    result.AddRange(request.CurrentPage
                        .Where(u => !IsMemberOfExceptionGroups(u.Id) && IsInactivePastThreshold(u, staleDate)));
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine($"Retrieval stoppd with {result.Count} entries: {ex.InnerException.Message}");
                }
            };
            return result;
        }
    }

    #endregion Iterators
}