using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    /// <summary>
    /// Exports country data into a NEW document created inside an already-running Office
    /// application instance (the add-in host), leaving it open and unsaved — no file is written.
    /// Implemented only by the NetOffice repositories. The host application is passed as
    /// <see cref="object"/> so this interface (and the BL manager that calls it) carry no
    /// NetOffice dependency; the implementation casts it to the concrete NetOffice type.
    /// </summary>
    public interface MsOfficeRunningAppExport__Repo__Interface
    {
        Task ExportToRunningAppAsync(
            object hostApplication,
            IList<MsOfficeCountry_Row> countries,
            Action<string> log);
    }
}
