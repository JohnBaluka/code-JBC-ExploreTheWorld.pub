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
    internal sealed class SimpleDbContextFactory<TContext> : IDbContextFactory<TContext>
        where TContext : DbContext
    {
        private readonly DbContextOptions<TContext> _options;

        public SimpleDbContextFactory(DbContextOptions<TContext> options) => _options = options;

        public TContext CreateDbContext() =>
            (TContext)Activator.CreateInstance(typeof(TContext), _options)!;
    }
}

// ---------------------------------------------------------------------------
// Polyfill: DbSet<T>.ExecuteDeleteAsync (introduced in EF Core 7.0)
// ---------------------------------------------------------------------------
namespace Microsoft.EntityFrameworkCore
{
}
