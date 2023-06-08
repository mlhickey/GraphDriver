//#define signinactivity
using Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace Services
{
    public class GuestUserServices
    {
        #region Privates

        private readonly GraphServiceClient _client;

        private readonly int _maxAttempts;

        private readonly int _staleRange;
        private readonly int _removalRange;
        private readonly int _staleInviteRange;

        private readonly List<string> _exemptionList;

        private Func<IDictionary<string, object>, string, string> GetDictionaryValue = (x, y) => x.ContainsKey(y) ? x[y].ToString() : "0";

        #endregion Privates

        public GuestUserServices(/*IGraphServiceClient graphServiceClient, IConfiguration configuration*/)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("local.settings.json")
                .Build();
            _client = new GraphServiceClientService(configuration).client;
            //
            // Set max values for retry and parallelism
            //
            _maxAttempts = configuration.GetValue("MaxAttempts", GuestUserConfiguration.MaxAttempts);
            //
            // Set removal range values to user staleness evaluation
            //
            var gracePeriod = SetValidRange(configuration.GetValue("GuestUser:GracePeriod", GuestUserConfiguration.GracePeriod), GuestUserConfiguration.GracePeriod);
            _staleRange = SetValidRange(configuration.GetValue("GuestUser:StaleRange", GuestUserConfiguration.StaleRange), GuestUserConfiguration.StaleRange);
            _staleInviteRange = SetValidRange(configuration.GetValue("GuestUserInvite:RemovalRange", GuestUserConfiguration.StaleRange), GuestUserConfiguration.StaleRange);

#if CanReadLeaverDate
            // Evaluated based on disable date stored in employeeLeaveDateTime
            _removalRange = gracePeriod;
#else
            // Evaluated based on lastSigninActivity
            _removalRange = _staleRange - gracePeriod;
#endif

            //
            // Check for exemption group, build list of value is non-null
            //
            var eGuid = configuration.GetValue("ExemptionGroupGUID", GuestUserConfiguration.ExemptionGroupGUID);
#if CheckMemberGroups
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

        private static string CurrentMethod([CallerMemberName] string name = "")
        {
            return name;
        }

        /// <summary>
        /// SetValidRange - ensures that associated range isn't below minimum threshold
        /// </summary>
        /// <param name="value"></param>
        /// <param name="baseline"></param>
        /// <returns>
        /// Valid int value for use with AddDays method
        /// </returns>
        private int SetValidRange(int value, int baseline)
        {
            if (value < baseline)
            {
                Console.WriteLine($"Stale range {value} is below minimum of {baseline}");
                value = baseline;
            }
            return value * -1;
        }

        #region ExemptUsers

        /// <summary>
        /// IsMemberOfExceptionGroups - check to see if specified id is a member of an exception group
        /// </summary>
        /// <param name="id"></param>
        /// <returns>bool</returns>
        public bool IsMemberOfExceptionGroups(string id)
        {
            if (_exemptionList.Count == 0)
            {
                Console.WriteLine("No exemption group specified, all users will be in scope");
                return false;
            }
#if CheckMemberGroups

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

#if !CheckMemberGroups

        /// <summary>
        /// GetExemptionGroupMembership - Creates user ID list basded on transitive membership
        /// </summary>
        /// <returns>List<string></returns>
        private async Task<List<string>> GetExemptionGroupMembership(string groupId)
        {
            var stopwatch = Stopwatch.StartNew();
            var exemptionList = new List<string>();

            try
            {
                var request = await _client.Groups[groupId]
                    .TransitiveMembers
                    .Request()
                    .Select("id")
                    .WithMaxRetry(_maxAttempts)
                    .Top(999)
                    .GetAsync();

                var pageIterator = PageIterator<DirectoryObject>.CreatePageIterator(_client, request, (u) =>
                {
                    exemptionList.Add(u.Id);
                    return true;
                });
                await pageIterator.IterateAsync();

                stopwatch.Stop();
                Console.WriteLine($"Total time: {stopwatch.Elapsed}");
                return exemptionList;
            }
            catch (ServiceException e) when (e.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                Console.WriteLine($"GetExemptionGroupMembership retrival failed: group {groupId} not found");
                throw e;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"GetExemptionGroupMembership retrival failed: {ex.InnerException.Message}");
                return new List<string>();
            }
        }

#endif

        #endregion ExemptUsers

        #region InactiveUsers

        /// <summary>
        /// GetDisableInactiveGuests - Retrieves list of disabled guest accounts which are outside removal threshold
        /// </summary>
        /// <returns>List<User>result</User></returns>
        public async Task<List<User>?> GetDisableInactiveGuests()
        {
            var caller = "program";
            var methodName = CurrentMethod();
            var staleDate = DateTime.Now.AddDays(_removalRange);
            var result = new List<User>();

            var queryBuilder = new[] {
                "accountEnabled eq false",
                "userType eq 'Guest'",
#if CanReadLeaverDate
                $"employeeLeaveDateTime le {staleDate:yyyy-MM-ddTHH:mm:ssZ}"
#endif
                };

            var queryFilter = string.Join(" and ", queryBuilder);
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("$count", "true")
            };

            Console.WriteLine($"GuestUserService::{methodName}:  Processing started at {DateTime.Now}");
#if DEBUG
            var stopWatch = new Stopwatch();
            stopWatch.Start();
#endif
            try
            {
                var request = await _client.Users
                   .Request(queryOptions)
                   .Header("ConsistencyLevel", "eventual")
                   .Filter(queryFilter)
                   .Select(string.Join(",", GuestUserConfiguration.Properties))
                   .WithMaxRetry(_maxAttempts)
                   .Top(999)
                   .GetAsync();
#if CanReadLeaverDate
                // Build list of validated users
                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    result.Add(u);
                    return true;
                });
#else
                // Build list of validated users based on inactivity evaluation
                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    if (IsInactivePastThreshold(u, staleDate))
                        result.Add(u);
                    return true;
                });
#endif
                await pageIterator.IterateAsync();
            }
            catch (ServiceException ex)
            {
                var failureType = result.Count > 0 ? "partially" : "completely";
                var message = ex?.InnerException?.Message ?? ex?.Message;
                Console.WriteLine($"{caller}::{methodName}:  retrieval {failureType} failed with {result?.Count} users: {message}");
            }
#if DEBUG
            stopWatch.Stop();
            Console.WriteLine($"{caller}::{methodName}: Query elapsed time: {stopWatch.Elapsed}");
            Console.WriteLine($"{caller}::{methodName}: Total users in scope:  {result?.Count} users");
#endif
            return result;
        }

        /// <summary>
        /// GetInactiveGuests - Retrieves list of active guest accounts which are:
        ///     enabled
        ///     outside staleness threshold
        ///     not members of the exemption group
        /// </summary>
        /// <returns></returns>
        public async Task<List<User>?> GetInactiveGuests()
        {
            var caller = "program";
            var methodName = CurrentMethod();
            var staleDate = DateTime.Now.AddDays(_staleRange);
            var result = new List<User>();
            int oos = 0;

#if signinactivity
            var queryFilter = $"signInActivity/lastSignInDateTime le {staleDate:yyyy-MM-ddTHH:mm:ssZ}";
#else
            // signInActivity queries are inconsistent due to scale, need to take an alternate approach
            // to return all guests of type enabled and evaluate separately

            var queryBuilder = new[] {
                "externalUserState eq 'Accepted'",
                "userType eq 'Guest'",
                "accountEnabled eq true"
                };
            var queryFilter = string.Join(" and ", queryBuilder);
#endif

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("$count", "true")
            };
            Console.WriteLine($"{caller}::{methodName}: Processing started at {DateTime.Now}");
#if DEBUG
            var stopWatch = new Stopwatch();
            stopWatch.Start();
#endif
            try
            {
                var request = await _client.Users
                   .Request(queryOptions)
                   .Filter(queryFilter)
#if !signinactivity
                   .Select(string.Join(",", GuestUserConfiguration.Properties))
#endif
                   .WithMaxRetry(_maxAttempts)
                   .Top(999)
                   .GetAsync();
#if signinactivity

                var pageIterator = PageIterator<User>.CreatePageIterator(_client, request, (u) =>
                {
                    if (u.UserType == "Guest" && u.ExternalUserState == "Accepted"
                        && u.AccountEnabled == true && !IsMemberOfExceptionGroups(u.Id))
                        result.Add(u);
                    return true;
                });
                await pageIterator.IterateAsync();
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
            }
            catch (ServiceException ex)
            {
                var failureType = result.Count > 0 ? "partially" : "completely";
                var message = ex?.InnerException?.Message ?? ex?.Message;
                Console.WriteLine($"{caller}::{methodName}:  retrieval {failureType} failed with {result?.Count} users: {message}");
            }
#if DEBUG
            stopWatch.Stop();
            Console.WriteLine($"{caller}::{methodName}: Query elapsed time: {stopWatch.Elapsed}");
            Console.WriteLine($"{caller}::{methodName}: Total users in scope:  {result?.Count - oos} users");
#endif
            return result;
        }

        #endregion InactiveUsers

        #region UnaccptedInvitations

        public async Task<List<User>?> GetUnacceptedInvitees([CallerMemberName] string caller = "")
        {
            var methodName = CurrentMethod();
            var staleDate = DateTime.Now.AddDays(_staleInviteRange);
            var result = new List<User>();

            var queryBuilder = new[] {
                "externalUserState eq 'PendingAcceptance'",
                "userType eq 'Guest'",
                $"createdDateTime le {staleDate:yyyy-MM-ddTHH:mm:ssZ}"
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
                   .Select(string.Join(",", GuestUserConfiguration.Properties))
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
                var message = ex?.InnerException?.Message ?? ex?.Message;
                Console.WriteLine($"{caller}::{methodName}:  retrieval {failureType} failed with {result?.Count} users: {message}");
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
            {
                if (user.SignInActivity?.LastSignInDateTime != null && user.SignInActivity?.LastNonInteractiveSignInDateTime != null)

                    return DateTimeOffset.Compare((DateTimeOffset)user.SignInActivity.LastSignInDateTime, (DateTimeOffset)user.SignInActivity.LastNonInteractiveSignInDateTime) > 0
                            ? user.SignInActivity.LastSignInDateTime
                            : user.SignInActivity.LastNonInteractiveSignInDateTime;
                else
                    return user.SignInActivity?.LastSignInDateTime ?? user.SignInActivity?.LastNonInteractiveSignInDateTime;
            }

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

}