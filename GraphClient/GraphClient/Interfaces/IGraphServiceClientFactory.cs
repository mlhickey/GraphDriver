using Microsoft.Graph;
using System.Threading.Tasks;

namespace Services
{
    public interface IGraphServiceClientFactory
    {
        Task<GraphServiceClient> CreateAsync();

        Task<GraphServiceClient> CreateAsync(string ver);

        Task<GraphServiceClient> CreateAsync(int timeOut);

        Task<GraphServiceClient> CreateAsync(string ver, int timeOut);
    }
}