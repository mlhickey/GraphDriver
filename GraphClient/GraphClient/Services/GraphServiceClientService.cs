using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Services;
using System;
using System.Net.Http.Headers;

namespace ConsoleApp.Services
{
    public class GraphServiceClientService : IGraphServiceClientService
    {
        private readonly string _graphEndpoint;
        private readonly string _graphUri;
        private readonly ITokenService _azureServiceTokenProvider = new TokenService();
        private readonly GraphServiceClient _graphServiceClient;
        private readonly int _maxAttempts;
        public GraphServiceClient client => _graphServiceClient;

        public GraphServiceClientService(IConfiguration configuration)
        {
            _maxAttempts = configuration.GetValue("MaxAttempts", 8);
            _graphEndpoint = $"https://graph.microsoft.{configuration.GetValue("AzureEnvironment", "com")}";
            _graphUri = $"{_graphEndpoint}/{configuration.GetValue("GraphVersion", "v1.0")}";

            _graphServiceClient = new GraphServiceClient(_graphUri,
                new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await _azureServiceTokenProvider.GetAccessTokenAsync(_graphEndpoint));
                    requestMessage.Headers.Add("ConsistencyLevel", "eventual");
                }));
            _graphServiceClient.HttpProvider.OverallTimeout = TimeSpan.FromMinutes(_maxAttempts);
        }
    }
}