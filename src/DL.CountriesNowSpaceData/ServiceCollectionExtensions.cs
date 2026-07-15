using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceData
{
    public static class ServiceCollectionExtensions
    {
        public static void EnsureExploreTheWorldDbCreated(this IServiceProvider serviceProvider)
        {
            using var scope = serviceProvider.CreateScope();
            var factory = scope.ServiceProvider
                .GetRequiredService<IDbContextFactory<ExploreTheWorldDbContext>>();
            using var context = factory.CreateDbContext();
            context.Database.EnsureCreated();
        }
    }
}
