using System.Threading.Tasks;

namespace Services
{
    public interface ITokenService
    {
        Task<string> GetAccessTokenAsync(string resource);
    }
}