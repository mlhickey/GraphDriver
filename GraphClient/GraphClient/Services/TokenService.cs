using Azure.Identity;
using System.Threading.Tasks;

namespace Services
{
    public class TokenService : ITokenService
    {
        private readonly DefaultAzureCredential _tokenCredential;

        public TokenService()
        {
#if DEBUG
            _tokenCredential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
            {
                ExcludeVisualStudioCredential = true,
                ExcludeInteractiveBrowserCredential = false
            });

#else
        _tokenCredential = new DefaultAzureCredential(new DefaultAzureCredentialOptions());
#endif
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
            var token = await _tokenCredential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { $"{resource}/.default" }));
            return token.Token;
        }
    }
}