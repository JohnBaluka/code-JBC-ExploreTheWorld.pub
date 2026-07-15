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
    internal static class DbSet_ExecuteDelete_Polyfill
    {
        internal static async Task<int> ExecuteDeleteAsync<TEntity>(
            this DbSet<TEntity> dbSet,
            CancellationToken cancellationToken = default)
            where TEntity : class
        {
            // DbSet<T> implements IInfrastructure<IServiceProvider> in EF Core 3.x
            var sp = ((IInfrastructure<IServiceProvider>)dbSet).Instance;
            var currentContext = (ICurrentDbContext)sp.GetService(typeof(ICurrentDbContext))!;
            var ctx = currentContext.Context;

            var entityType = ctx.Model.FindEntityType(typeof(TEntity))
                ?? throw new InvalidOperationException(
                    $"Entity type '{typeof(TEntity).Name}' not found in the EF Core model.");

            var tableName = entityType.GetTableName()!;
            var schema = entityType.GetSchema();
            var fullName = schema == null
                ? $"[{tableName}]"
                : $"[{schema}].[{tableName}]";

            return await ctx.Database.ExecuteSqlRawAsync(
                $"DELETE FROM {fullName}", cancellationToken);
        }
    }
}
