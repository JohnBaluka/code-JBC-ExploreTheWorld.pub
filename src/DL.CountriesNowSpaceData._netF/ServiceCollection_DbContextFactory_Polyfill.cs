// EF Core 3.1 polyfills for APIs introduced in later EF Core versions.
// This file is compiled only by the DL.CountriesNowSpaceData._netF project.

using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;

// ---------------------------------------------------------------------------
// Polyfill: IServiceCollection.AddDbContextFactory<T> (introduced in EF Core 5.0)
// ---------------------------------------------------------------------------
namespace Microsoft.Extensions.DependencyInjection
{
    internal static class ServiceCollection_DbContextFactory_Polyfill
    {
        internal static IServiceCollection AddDbContextFactory<TContext>(
            this IServiceCollection services,
            Action<DbContextOptionsBuilder> optionsAction,
            ServiceLifetime lifetime = ServiceLifetime.Singleton)
            where TContext : DbContext
        {
            var builder = new DbContextOptionsBuilder<TContext>();
            optionsAction(builder);
            var options = builder.Options;

            services.Add(new ServiceDescriptor(
                typeof(IDbContextFactory<TContext>),
                sp => (object)new SimpleDbContextFactory<TContext>(options),
                lifetime));

            return services;
        }
    }
}
