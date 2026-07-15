using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;

namespace JBC.ExploreTheWorld.AL
{
    /// <summary>
    /// Produces Office documents (Word/Excel/PowerPoint) from in-memory data.
    /// Implemented by the host layer so Blazor pages remain platform-neutral.
    /// </summary>
    public interface OfficeExport_AppService__Interface
    {
        /// <summary>
        /// The Office-library options this host actually supports, in display order. Drives the
        /// Countries page "Library" picklist. Browser hosts (OpenXML-only in the WASM sandbox)
        /// return just "OpenXML"; desktop hosts return Interop/Dynamic/NetOffice/OpenXML.
        /// </summary>
        IReadOnlyList<string> SupportedLibraries { get; }

        /// <summary>
        /// True when the export writes to the supplied local file path (desktop hosts); false when
        /// it triggers a browser download by file name only (browser/WASM hosts, which have no
        /// local file system). Drives whether the Countries page shows a "File Path" or a
        /// "File Name" field and whether the default includes a Documents-folder path.
        /// </summary>
        bool SavesToFilePath { get; }

        Task ExportCountriesAsync(
            IList<CountryBasic_Row> countries,
            string exportType,
            string exportLibrary,
            string filePath,
            Action<string> log);
    }
}
