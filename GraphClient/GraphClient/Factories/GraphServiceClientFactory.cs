using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Services
{
    public class GraphServiceClientFactory : IGraphServiceClientFactory
    {
        private readonly string _graphEndpoint;
        private readonly string _graphUri;
        private readonly ITokenService _azureServiceTokenProvider;
        private readonly int _maxAttempts;

        public GraphServiceClientFactory()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("local.settings.json")
                .Build();

            _azureServiceTokenProvider = new TokenService();
            _maxAttempts = configuration.GetValue<int>("MaxAttempts", 8);
            _graphEndpoint = $"https://graph.microsoft.{configuration.GetValue("AzureEnvironment", "com")}";
            _graphUri = $"{_graphEndpoint}/{configuration.GetValue("GraphVersion", "v1.0")}";
        }

        public Task<GraphServiceClient> CreateAsync()
        {
            return CreateAsync(_graphUri, _maxAttempts);
        }

        public Task<GraphServiceClient> CreateAsync(string graphVersion)
        {
            return CreateAsync($"{_graphEndpoint}/{graphVersion}", _maxAttempts);
        }

        public Task<GraphServiceClient> CreateAsync(int timeOut)
        {
            return CreateAsync(_graphUri, timeOut);
        }

        /// <summary>
        /// CreateAsync - create new GraphServceClient using specified Uri and timeout period
        /// </summary>
        /// <param name="Uri"></param>
        /// <param name="timeOut"></param>
        /// <returns></returns>
        public Task<GraphServiceClient> CreateAsync(string Uri, int timeOut)
        {
            GraphServiceClient graphClient = new GraphServiceClient(Uri,
                new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await _azureServiceTokenProvider.GetAccessTokenAsync(_graphEndpoint));
                    requestMessage.Headers.Add("ConsistencyLevel", "eventual");
                }));
            graphClient.HttpProvider.OverallTimeout = TimeSpan.FromMinutes(timeOut);
            return Task.FromResult(graphClient);
        }
    }
}