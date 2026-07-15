using System.Collections.Generic;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    internal static class ExportMenuHelper
    {
        internal static void RunExport(ExportType type, ExportMethod method,
            IList<MsOfficeCountry_Row> countries, MsOfficeExportManager__Service exportManager)
            => new ExportLog_Form(type, method, countries, exportManager).Show();
    }
}
