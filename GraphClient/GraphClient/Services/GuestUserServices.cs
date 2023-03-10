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
        private readonly IGraphServiceClientFactory _clientFactory;
        private readonly List<string> _exemptionList;

        private readonly int _staleRange;
        private readonly int _staleInviteRange;
        private readonly int _removalRange;
        private readonly string _guestProperties;

        private Func<IDictionary<string, object>, string, string> GetDictionaryValue = (x, y) => (string)(x.ContainsKey(y) ? x[y].ToString() : "0");

        public GuestUserServices()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("local.settings.json")
                .Build();
            _clientFactory = new GraphServiceClientFactory();

            _guestProperties = configuration.GetValue<string>("GuestUserProps", GuestUserConfiguration.StaleGuestProperties);
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
            if (!string.IsNullOrEmpty(eGuid))
            {
                this._exemptionList = GetExemptionGroupMembership(eGuid).Result;
            }
            else
            {
                _exemptionList = new List<string>();
            }
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

        /// <summary>
        /// GetExemptionGroupGUIDMembership - retrieves membership of specifed group GUID
        /// </summary>
        /// <returns>
        /// Collection of group members
        /// </returns>
        private async Task<List<string>> GetExemptionGroupMembership(string exemptionGUID)
        {
            var client = await _clientFactory.CreateAsync();
            var exemptionMembers = await client.Groups[exemptionGUID]
                .Members
                .Request()
                .Select("id")
                .GetAsync();

            var exmptionList = exemptionMembers
                .Select(i => i.Id).ToList();
            return exmptionList;
        }

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
            var client = await _clientFactory.CreateAsync();
            var isEnabled = enabled.ToString().ToLower();

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
            try
            {
                var request = await client.Users
                   .Request(queryOptions)
                   .Filter(string.Join(" and ", queryBuilder))
                   .Select(_guestProperties)
                   .Top(999)
                   .GetAsync();
                Console.WriteLine($"{callerName}::GetGuestAccounts: Total of {GetDictionaryValue(request.AdditionalData, "@odata.count")} objects");
                result = await ProcessBoundRequestList(request, staleDate);
            }
            catch (Exception ex)
            {
                var failureType = result.Count() > 0 ? "partially" : "completely";
                Console.WriteLine($"GetInactiveGuests retrival failed {failureType}: {ex.Message}");
            }
            return result;
        }

        #endregion InactiveUsers

        #region UnaccptedInvitations

        public async Task<List<User>> GetUnacceptedInvitees()
        {
            var result = new List<User>();
            var client = await _clientFactory.CreateAsync();
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
                var request = await client.Users
                   .Request(queryOptions)
                   .Filter(string.Join(" and ", queryBuilder))
                   .Select(_guestProperties)
                   .Top(999)
                   .GetAsync();
                Console.WriteLine($"{GetType().Name}: Total of {GetDictionaryValue(request.AdditionalData, "@odata.count")} objects");
                result = await ProcessRequestList(request);
                return result;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"GetUnacceptedInvitees retrival failed: {ex.Message}");
                return new List<User>();
            }
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
                return user.SignInActivity.LastSignInDateTime > user.SignInActivity.LastNonInteractiveSignInDateTime
                        ? user.SignInActivity.LastSignInDateTime
                        : user.SignInActivity.LastNonInteractiveSignInDateTime;

            if (user.CreatedDateTime != null && user.SignInSessionsValidFromDateTime != null)
            {
                return user.CreatedDateTime > user.SignInSessionsValidFromDateTime
                        ? user.CreatedDateTime
                        : user.SignInSessionsValidFromDateTime;
            }
            return user.CreatedDateTime ?? user.SignInSessionsValidFromDateTime;
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

            while (request != null)
            {
                result.AddRange(request);
                if (request.NextPageRequest == null) break;
                request = await request.NextPageRequest
                    .GetAsync();
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

            while (request != null)
            {
                result.AddRange(request
                    .Where(u => !_exemptionList.Contains(u.Id) && IsInactivePastThreshold(u, staleDate))
                    .ToList());
                if (request.NextPageRequest == null) break;
                request = await request.NextPageRequest
                    .GetAsync();
            };
            return result;
        }
    }
}