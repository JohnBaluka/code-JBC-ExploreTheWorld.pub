using System.Threading;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace Microsoft.EntityFrameworkCore
{
    /// <summary>
    /// Polyfill for IDbContextFactory&lt;TContext&gt; which was introduced in EF Core 5.0.
    /// EF Core 3.1 (the last version supporting .NET Framework 4.8.1) does not include this interface.
    /// </summary>
    public interface IDbContextFactory<TContext> where TContext : DbContext
    {
        TContext CreateDbContext();
    }
}
