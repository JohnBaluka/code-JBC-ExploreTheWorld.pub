using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// WinForms host implementation of <see cref="OfficeExport_AppService__Interface"/>. Maps the
    /// Blazor-facing <c>CountryBasic_Row</c> data to <c>MsOfficeCountry_Row</c> and delegates the
    /// export to the business-layer <see cref="MsOfficeExportManager__Service"/> — it never
    /// instantiates DL export repositories directly.
    /// </summary>
    public class OfficeExport_AppService : OfficeExport_AppService__Interface
    {
        private readonly MsOfficeExportManager__Service _exportManager;

        // Desktop hosts have a local Office install, so all export libraries are available.
        public IReadOnlyList<string> SupportedLibraries { get; } =
            new[] { "Interop", "Dynamic", "NetOffice", "OpenXML" };

        // Desktop hosts write the document to the local file path supplied by the user.
        public bool SavesToFilePath => true;

        public OfficeExport_AppService(MsOfficeExportManager__Service exportManager)
        {
            _exportManager = exportManager;
        }

        public Task ExportCountriesAsync(
            IList<CountryBasic_Row> countries,
            string exportType,
            string exportLibrary,
            string filePath,
            Action<string> log)
        {
            var rows = countries
                .Select(c => new MsOfficeCountry_Row(c.Country, c.Iso2, c.Iso3))
                .ToList();

            return _exportManager.ExportAsync(rows, exportType, exportLibrary, filePath, log);
        }
    }
}
