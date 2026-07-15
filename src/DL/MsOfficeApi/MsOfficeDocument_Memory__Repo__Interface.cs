using System;
using System.Collections.Generic;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    /// <summary>
    /// Builds an Office document entirely in memory (no local file system) and returns its bytes,
    /// for hosts that stream a browser download instead of writing to a path — i.e. the Blazor
    /// (WASM) browser export. A DL contract implemented by the OpenXML `_Impl` project and
    /// registered by the browser host applications, so the AL Blazor library never references the
    /// platform-specific DL repo `_Impl` project directly.
    /// </summary>
    public interface MsOfficeDocument_Memory__Repo__Interface
    {
        byte[] BuildWordBytes(IList<MsOfficeCountry_Row> countries, Action<string> log);
        byte[] BuildExcelBytes(IList<MsOfficeCountry_Row> countries, Action<string> log);
        byte[] BuildPowerPointBytes(IList<MsOfficeCountry_Row> countries, Action<string> log);
    }
}
