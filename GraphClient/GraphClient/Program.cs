using Microsoft.Graph;
using Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp
{
    internal delegate Task<List<User>> InactiveUserServices();

    internal class Program
    {
        private static GuestUserServices _guestUserServices;
        private static List<InactiveUserServices> _services;

        private static async Task Main(string[] args)
        {
            _guestUserServices = new GuestUserServices();
            _services = new List<InactiveUserServices>
            {
                //_guestUserServices.GetUnacceptedInvitees,
                _guestUserServices.GetDisableInactiveGuests,
                _guestUserServices.GetInactiveGuests
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
#if !NET6_0
            foreach (var s in _services)
#else
            await Parallel.ForEachAsync(_services, async (s, cancellationToken) =>
#endif
            {
                Console.WriteLine($"{s.GetType().Name}::{s.Method.Name} start");

                var threadWatch = new Stopwatch();
                threadWatch.Start();
                var uRet = await s.Invoke();
                threadWatch.Stop();
                Console.WriteLine($"{uRet.Count} users returned for {s.Method.Name}");

#if DEBUG
                int exempt = 0;
                ParallelOptions parallelOptions = new()
                {
                    MaxDegreeOfParallelism = 100
                };
                await Parallel.ForEachAsync(uRet, async (u, cancellationToken) =>
                {
                    if (await _guestUserServices.IsMemberOfExceptionGroups(u.Id))
                        exempt++;
                });
                //var eRet = uRet.Where(u => !_guestUserServices.IsMemberOfExceptionGroups(u.Id).Result).ToList();
                //var exempt = uRet.Count - eRet.Count;
                Console.WriteLine($"Total of {exempt} users exempted");
                using (StreamWriter writer = new StreamWriter(new FileStream($"{s.Method.Name}.csv", FileMode.Create, FileAccess.Write)))
                {
                    writer.WriteLine("displayName,id,AccountEnabled");
                    foreach (var u in uRet)
                    {
                        /*
                        var lastLogin = _guestUserServices.GetLastSignIn(u).Value.ToString("yyyy-MM-dd");
                        writer.WriteLine($"{u.UserPrincipalName},{u.Id},{lastLogin}");
                        */
                        writer.WriteLine($"{u.UserPrincipalName},{u.Id},{u.AccountEnabled}");
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