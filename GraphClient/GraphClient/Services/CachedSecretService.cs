using Microsoft.Extensions.Caching.Memory;
using System;
using System.Runtime.InteropServices;
using System.Security;

namespace SyncProvisioning.Services
{
    public class CachedSecretService
    {
        private readonly IMemoryCache _memoryCache;

        public CachedSecretService(IMemoryCache memoryCache)
        {
            _memoryCache = memoryCache;
        }

        public string? GetAccessToken(string name)
        {
            var existing = _memoryCache.Get<SecureString>(name);
            return existing != null ? ConvertFromSecureString(existing) : null;
        }

        public void StoreAccessToken(string name, string secret)
        {
            var expirationTimespan = TimeSpan.FromMinutes(15);
            _memoryCache.Set(name, ConvertToSecureString(secret), expirationTimespan);
        }

        private SecureString ConvertToSecureString(string value)
        {
            var secureString = new SecureString();
            foreach (var character in value.ToCharArray())
            {
                secureString.AppendChar(character);
            }

            return secureString;
        }

        private string ConvertFromSecureString(SecureString value)
        {
            var valuePtr = IntPtr.Zero;
            try
            {
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(value);
                return Marshal.PtrToStringUni(valuePtr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }
    }
}