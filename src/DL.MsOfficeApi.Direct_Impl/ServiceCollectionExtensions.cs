using JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddDirectLib(this IServiceCollection services)
        {
            services.AddSingleton<MsWord__Repo__Interface,        MsWord_Direct__Repo>();
            services.AddSingleton<MsExcel__Repo__Interface,       MsExcel_Direct__Repo>();
            services.AddSingleton<MsPowerPoint__Repo__Interface,  MsPowerPoint_Direct__Repo>();
            return services;
        }
    }
}
