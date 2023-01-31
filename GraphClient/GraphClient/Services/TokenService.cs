using Azure.Identity;
using System.Threading.Tasks;

namespace Services
{
    public class TokenService : ITokenService
    {
        private readonly CachedSecretService _memoryCache;
#if DEBUG

        private readonly DefaultAzureCredential _tokenCredential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
        {
            // VisualStudioCredential requires login to desired tenant within current VS session
            // If a different account is required for target, set to true
            ExcludeVisualStudioCredential = false,
            // Exclude to avoid CLI OBO deviceCode failure due to CA policies
            ExcludeAzureCliCredential = true,
            ExcludeInteractiveBrowserCredential = false
        });

#else
        private readonly DefaultAzureCredential _tokenCredential = new DefaultAzureCredential(new DefaultAzureCredentialOptions());
#endif

        public TokenService()
        {
            _memoryCache = new CachedSecretService();
        }

        /// <summary>
        /// GetAccessTokenAsync - retrieve access token for specified resource:
        ///     cached credential if cached
        ///     new credential if not cached
        /// </summary>
        /// <param name="resource"></param>
        /// <returns>>Access token as string</returns>
        public async Task<string> GetAccessTokenAsync(string resource)
        {
            var existing = _memoryCache.GetAccessToken(resource);
            if (existing != null) return existing;

            var token = await _tokenCredential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { $"{resource}/.default" }));

            _memoryCache.StoreAccessToken(resource, token.Token);
            return token.Token;
        }
    }
}