using JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddDynamicLib(this IServiceCollection services)
        {
            services.AddSingleton<MsWord__Repo__Interface,        MsWord_Dynamic__Repo>();
            services.AddSingleton<MsExcel__Repo__Interface,       MsExcel_Dynamic__Repo>();
            services.AddSingleton<MsPowerPoint__Repo__Interface,  MsPowerPoint_Dynamic__Repo>();
            return services;
        }
    }
}
