using System.Threading;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace Microsoft.EntityFrameworkCore
{
    /// <summary>
    /// Polyfill for CreateDbContextAsync extension method which was introduced in EF Core 5.0.
    /// </summary>
    public static class DbContextFactoryExtensions
    {
        public static Task<TContext> CreateDbContextAsync<TContext>(this IDbContextFactory<TContext> factory, CancellationToken cancellationToken = default)
            where TContext : DbContext
        {
            return Task.FromResult(factory.CreateDbContext());
        }
    }
}
