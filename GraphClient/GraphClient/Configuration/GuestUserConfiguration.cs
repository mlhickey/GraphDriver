namespace Configuration
{
    internal static class GuestUserConfiguration
    {
        public const int StaleRange = 60;
        public const int StaleInviteRange = 90;
        public const int RemovalRange = 74;
        public const int MaxParallel = 100;
        public const string InviteeProperties = "id,displayName,userPrincipalName,externalUserState,createdDateTime,signInSessionsValidFromDateTime";
        public const string StaleGuestProperties = InviteeProperties + ",signInActivity";
        public const string GraphVersion = "beta";

        // dsr-guest-lcm-exception
        public const string ExemptionGroupGUID = "adaadda0-1462-435e-8115-cba7d2fd6d73";
    }
}