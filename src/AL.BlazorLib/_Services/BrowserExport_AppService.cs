using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsJSInterop.FileDownload;

namespace JBC.ExploreTheWorld.AL.BlazorLib
{
    // Implements OfficeExport_AppService__Interface for browser (WASM). Builds the document in
    // memory via the host-supplied MsOfficeDocument_Memory__Repo__Interface (an OpenXML impl the
    // browser hosts register) and triggers a browser download via the FileDownload JS interop.
    // When a FlagImageManager is registered, rows are enriched with cached/downloaded flag PNGs.
    public class BrowserExport_AppService : OfficeExport_AppService__Interface
    {
        private readonly FileDownload__Interop__Interface _fileDownload;
        private readonly FlagImageManager__Service? _flagImageManager;
        private readonly MsOfficeDocument_Memory__Repo__Interface _repo;

        // OpenXML is the only library available in the browser (WASM) sandbox — Interop/Dynamic/
        // NetOffice need a local Office COM install, so they are not offered on browser hosts.
        public IReadOnlyList<string> SupportedLibraries { get; } = new[] { "OpenXML" };

        // Browser hosts have no local file system — the document is streamed to the browser as a
        // download (by file name), not written to a path on disk.
        public bool SavesToFilePath => false;

        public BrowserExport_AppService(
            MsOfficeDocument_Memory__Repo__Interface repo,
            FileDownload__Interop__Interface fileDownload,
            FlagImageManager__Service? flagImageManager = null)
        {
            _repo             = repo;
            _fileDownload     = fileDownload;
            _flagImageManager = flagImageManager;
        }

        public async Task ExportCountriesAsync(
            IList<CountryBasic_Row> countries,
            string exportType,
            string exportLibrary,
            string filePath,
            Action<string> log)
        {
            if (exportLibrary is "NetOffice" or "Interop" or "Dynamic")
                log($"{exportLibrary} is not available in the browser. Switching to OpenXml.");

            var fileName = Path.GetFileName(filePath);
            if (string.IsNullOrEmpty(fileName))
                fileName = JBC.ExploreTheWorld.CL.MsOfficeExportName_Helper.BuildFileName(
                    "ETW_CountriesNow", exportLibrary, fromAccessDb: false, GetExtension(exportType));

            var mapped = countries
                .Select(c => new MsOfficeCountry_Row(c.Country ?? "", c.Iso2 ?? "", c.Iso3 ?? ""))
                .ToList();

            if (_flagImageManager != null)
            {
                log($"Resolving flag images for {mapped.Count} countries (local cache first)...");
                var images = await _flagImageManager.GetFlagImagesAsync(mapped.Select(c => c.Iso2), log);
                mapped = mapped
                    .Select(c => images.TryGetValue(c.Iso2, out var image)
                        ? c with { FlagPng = image.ImageBytes }
                        : c)
                    .ToList();
            }

            byte[] bytes;
            string contentType;

            switch (exportType)
            {
                case "Excel":
                    bytes       = _repo.BuildExcelBytes(mapped, log);
                    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    break;
                case "PowerPoint":
                    bytes       = _repo.BuildPowerPointBytes(mapped, log);
                    contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    break;
                default:
                    bytes       = _repo.BuildWordBytes(mapped, log);
                    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    break;
            }

            log("Triggering browser download...");
            await _fileDownload.DownloadFileFromBytesAsync(fileName, contentType, bytes);
            log($"✔ Download triggered: {fileName} ({bytes.Length:N0} bytes)");
        }

        private static string GetExtension(string exportType) => exportType switch
        {
            "Excel"      => "xlsx",
            "PowerPoint" => "pptx",
            _            => "docx"
        };
    }
}
