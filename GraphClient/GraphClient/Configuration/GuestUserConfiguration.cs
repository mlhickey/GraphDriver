namespace Configuration
{
    internal static class GuestUserConfiguration
    {
        // Maximum idle time before disablement
        public const int StaleRange = 60;

        // Time from disablement to deletion
        public const int GracePeriod = 14;

        // Maximum time for unaccepted invites before removal
        public const int StaleInviteRange = 90;

        public const int MaxParallel = 100;
        public const int MaxAttempts = 10;

        //public const string Properties = "id,displayName,userPrincipalName,externalUserState,createdDateTime,signInSessionsValidFromDateTime,signInActivity";
        public static string[] Properties = {
                "id",
                "displayName",
                "userPrincipalName",
                "externalUserState",
                "createdDateTime",
                "signInSessionsValidFromDateTime",
                "signInActivity"
            };

        public const string GraphVersion = "beta";

        // GUID of group used for exemtption from idle user processing
        // Single group with transitive membership checks
        public const string ExemptionGroupGUID = "adaadda0-1462-435e-8115-cba7d2fd6d73";
    }
}