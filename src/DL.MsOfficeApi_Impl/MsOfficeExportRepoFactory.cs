using System;
using JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    /// <summary>
    /// Host-side factory that instantiates the concrete Office document repositories. This is the
    /// single place permitted to reference the platform-specific DL repo projects; UI items obtain
    /// the repository indirectly through <see cref="MsOfficeExportManager__Service"/>.
    /// </summary>
    public class MsOfficeExportRepoFactory : MsOfficeExportRepoFactory__Interface
    {
        public MsWord__Repo__Interface GetWordRepo(string exportMethod) => exportMethod switch
        {
            "Direct"    => new MsWord_Direct__Repo(),
            "Interop"   => new MsWord_Interop__Repo(),
            "Dynamic"   => new MsWord_Dynamic__Repo(),
            "NetOffice" => new MsWord_NetOffice__Repo(),
            _           => new MsWord_OpenXml__Repo()
        };

        public MsExcel__Repo__Interface GetExcelRepo(string exportMethod) => exportMethod switch
        {
            "Direct"    => new MsExcel_Direct__Repo(),
            "Interop"   => new MsExcel_Interop__Repo(),
            "Dynamic"   => new MsExcel_Dynamic__Repo(),
            "NetOffice" => new MsExcel_NetOffice__Repo(),
            _           => new MsExcel_OpenXml__Repo()
        };

        public MsPowerPoint__Repo__Interface GetPowerPointRepo(string exportMethod) => exportMethod switch
        {
            "Direct"    => new MsPowerPoint_Direct__Repo(),
            "Interop"   => new MsPowerPoint_Interop__Repo(),
            "Dynamic"   => new MsPowerPoint_Dynamic__Repo(),
            "NetOffice" => new MsPowerPoint_NetOffice__Repo(),
            _           => new MsPowerPoint_OpenXml__Repo()
        };
    }
}
