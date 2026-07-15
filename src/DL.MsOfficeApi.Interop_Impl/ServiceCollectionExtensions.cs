using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddInteropLib(this IServiceCollection services)
        {
            services.AddSingleton<MsWord__Repo__Interface,        MsWord_Interop__Repo>();
            services.AddSingleton<MsExcel__Repo__Interface,       MsExcel_Interop__Repo>();
            services.AddSingleton<MsPowerPoint__Repo__Interface,  MsPowerPoint_Interop__Repo>();
            return services;
        }
    }
}
