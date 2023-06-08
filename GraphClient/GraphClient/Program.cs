using Microsoft.Graph;
using Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

/*
using IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddSingleton<GraphServiceClientService>();
        services.AddSingleton<ITokenService, TokenService>();
        services.AddSingleton<GraphServiceClient>();
    })
    .Build();
*/

namespace ConsoleApp
{
    internal delegate Task<List<User>> InactiveUserServices();

    internal class Program
    {
        private static GuestUserServices _guestUserServices = new GuestUserServices();
        private static List<InactiveUserServices> _services;

        private static async Task Main(string[] args)
        {
            _services = new List<InactiveUserServices>
            {
                //_guestUserServices.GetUnacceptedInvitees,
                _guestUserServices.GetDisableInactiveGuests
                //_guestUserServices.GetInactiveGuests
            };

            await GraphDriver();
        }

        /// <summary>
        /// Call MS Graph and print results
        /// </summary>
        /// <returns></returns>
        private static async Task GraphDriver()
        {
            Stopwatch stopWatch = new Stopwatch(); ;
            stopWatch.Start();
#if NET6_0
            await Parallel.ForEachAsync(_services, async (s, cancellationToken) =>
#else
            foreach (var s in _services)
#endif
            {
                Console.WriteLine($"{s.GetType().Name}::{s.Method.Name} start");

                var threadWatch = new Stopwatch();
                threadWatch.Start();
                var uRet = await s.Invoke();
                threadWatch.Stop();
#if DEBUG
                Console.WriteLine($"{uRet.Count} users returned for {s.Method.Name}");
                using (StreamWriter writer = new StreamWriter(new FileStream($"{s.Method.Name}.csv", FileMode.Create, FileAccess.Write)))
                {
                    writer.WriteLine("displayName,id,AccountEnabled,lastSignin");
                    foreach (var u in uRet)
                    {
                        var last = _guestUserServices.GetLastSignIn(u).Value.DateTime;
                        writer.WriteLine($"{u?.UserPrincipalName},{u?.Id},{u?.AccountEnabled},{last.ToString("yyyy-MM-dd")}");
                    }
                }
#endif
                Console.WriteLine($"{s.GetType().Name}::{s.Method.Name} - {uRet.Count} objects in scope, elapsed time {threadWatch.Elapsed}");
            }
#if NET6_0
            );
#endif
            stopWatch.Stop();
            Console.WriteLine($"All methods complete, overall runtime of {stopWatch.Elapsed}");
        }
    }
}