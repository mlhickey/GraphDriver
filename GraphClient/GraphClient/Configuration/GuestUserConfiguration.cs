namespace Configuration
{
    internal static class GuestUserConfiguration
    {
        public const int StaleRange = 90;
        public const int StaleInviteRange = 90;
        public const int RemovalRange = 180;
        public const int MaxParallel = 100;
        public const string InviteeProperties = "id,displayName,userPrincipalName,externalUserState,createdDateTime,signInSessionsValidFromDateTime";
        public const string StaleGuestProperties = InviteeProperties + ",signInActivity";
        public const string GraphVersion = "beta";
        // dsr-guest-lcm-exception
        public const string ExemptionGroupGUID = "e6baf2ff-d76b-4608-9239-93daf30dfbee";
    }
}