using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddNetOfficeLib(this IServiceCollection services)
        {
            services.AddSingleton<MsWord__Repo__Interface,        MsWord_NetOffice__Repo>();
            services.AddSingleton<MsExcel__Repo__Interface,       MsExcel_NetOffice__Repo>();
            services.AddSingleton<MsPowerPoint__Repo__Interface,  MsPowerPoint_NetOffice__Repo>();
            return services;
        }
    }
}
