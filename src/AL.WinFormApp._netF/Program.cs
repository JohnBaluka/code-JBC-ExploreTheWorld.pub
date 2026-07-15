using System;
using System.Configuration;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using System.Windows.Forms;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.AccessDb_Impl;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqlServerDb_Impl;

namespace JBC.ExploreTheWorld.AL.WinFormApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var dbProviderName = ConfigurationManager.AppSettings["DbProvider"] ?? "SqlServerDb";

            ExploreTheWorldDbContextFactory dbFactory;
            switch (dbProviderName)
            {
                case "AccessDb":
                    var accessDbPath = ConfigurationManager.ConnectionStrings["AccessDb"]?.ConnectionString
                        ?? ExploreTheWorldAccessDb.DefaultDbPath;
                    dbFactory = ExploreTheWorldAccessDb.CreateFactory(accessDbPath);
                    break;

                default: // "SqlServerDb"
                    var sqlServerConnStr = ConfigurationManager.ConnectionStrings["SqlServerDb"]?.ConnectionString
                        ?? ExploreTheWorldSqlServerDb.DefaultConnectionString;
                    dbFactory = ExploreTheWorldSqlServerDb.CreateFactory(sqlServerConnStr);
                    break;
            }

            JBC.ExploreTheWorld.DL.CountriesNowSpaceData.ServiceCollectionExtensions.EnsureExploreTheWorldDbCreated(dbFactory);

            var countriesNowApiService = new CountriesNowSpaceApi__Repo();
            var countriesNowDbManager = new CountriesNowSpaceApiManager__Repo(dbFactory);

            CountriesNowSpaceManager__Service countriesNowManager =
                new CountriesNowSpaceManager__Service(countriesNowApiService, countriesNowDbManager);

            // Supply the platform-specific Office composition (export factory + Save-As-JSON writer)
            // to the WinForms UI libraries, which reference no DL repo _Impl project.
            MsOfficeSaveAsJsonWriterProvider.Current  = new MsOfficeSaveAsJsonWriter();
            MsOfficeExportRepoFactoryProvider.Current = new MsOfficeExportRepoFactory();

            Application.Run(new Main_Form(countriesNowManager));
        }
    }
}
