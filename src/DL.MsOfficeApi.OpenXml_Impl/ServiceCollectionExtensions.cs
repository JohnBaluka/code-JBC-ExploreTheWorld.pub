using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using Microsoft.Extensions.DependencyInjection;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddOpenXmlLib(this IServiceCollection services)
        {
            services.AddSingleton<MsWord__Repo__Interface,        MsWord_OpenXml__Repo>();
            services.AddSingleton<MsExcel__Repo__Interface,       MsExcel_OpenXml__Repo>();
            services.AddSingleton<MsPowerPoint__Repo__Interface,  MsPowerPoint_OpenXml__Repo>();
            return services;
        }
    }
}
