using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Net.Http.Headers;

namespace Services
{
    public class GraphServiceClientService : IGraphServiceClientService
    {
        private readonly ITokenService tokenService = new TokenService();
        private readonly GraphServiceClient _graphServiceClient;
        public GraphServiceClient client => _graphServiceClient;

        public GraphServiceClientService(IConfiguration configuration /*,ITokenService tokenService*/)
        {
            var maxAttempts = configuration.GetValue("MaxAttempts", 8);
            var graphEndpoint = $"https://graph.microsoft.{configuration.GetValue("AzureEnvironment", "com")}";
            var graphUri = $"{graphEndpoint}/{configuration.GetValue("GraphVersion", "v1.0")}";

            _graphServiceClient = new GraphServiceClient(graphUri,
                new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await tokenService.GetAccessTokenAsync(graphEndpoint));
                }));
            _graphServiceClient.HttpProvider.OverallTimeout = TimeSpan.FromMinutes(maxAttempts);
        }
    }
}