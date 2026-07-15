using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.BL
{
    /// <summary>
    /// Selects the target Office host (Word/Excel/PowerPoint) and delegates to the concrete
    /// DL repository resolved by <see cref="MsOfficeExportRepoFactory__Interface"/>. Holds no
    /// reference to the platform-specific DL repo types, so it remains framework-neutral.
    /// When a <see cref="FlagImageManager__Service"/> is supplied, country rows are
    /// enriched with cached/downloaded PNG flag images before they reach the writers.
    /// </summary>
    public class MsOfficeExportManager__Service
    {
        private readonly MsOfficeExportRepoFactory__Interface _factory;
        private readonly FlagImageManager__Service? _flagImageManager;

        public MsOfficeExportManager__Service(
            MsOfficeExportRepoFactory__Interface factory,
            FlagImageManager__Service? flagImageManager = null)
        {
            _factory          = factory;
            _flagImageManager = flagImageManager;
        }

        public async Task ExportAsync(
            IList<MsOfficeCountry_Row> countries,
            string exportType,
            string exportMethod,
            string filePath,
            Action<string> log)
        {
            var rows = await EnrichWithFlagImagesAsync(countries, log);

            await (exportType switch
            {
                "Excel"      => _factory.GetExcelRepo(exportMethod).ExportAsync(rows, filePath, log),
                "PowerPoint" => _factory.GetPowerPointRepo(exportMethod).ExportAsync(rows, filePath, log),
                "Word"       => _factory.GetWordRepo(exportMethod).ExportAsync(rows, filePath, log),
                _            => Task.FromException(new NotSupportedException(
                                    $"Unsupported export type '{exportType}'."))
            });
        }

        public async Task ExportToRunningAppAsync(
            string exportType,
            object hostApplication,
            IList<MsOfficeCountry_Row> countries,
            Action<string> log)
        {
            // Running-app export always uses the NetOffice repository (COM automation of the
            // live host). Only the NetOffice repos implement MsOfficeRunningAppExport__Repo__Interface.
            object? repo = exportType switch
            {
                "Excel"      => _factory.GetExcelRepo("NetOffice"),
                "PowerPoint" => _factory.GetPowerPointRepo("NetOffice"),
                "Word"       => _factory.GetWordRepo("NetOffice"),
                _            => null
            };

            if (repo is not MsOfficeRunningAppExport__Repo__Interface runningAppRepo)
                throw new NotSupportedException(
                    $"Running-app export is not supported for type '{exportType}'.");

            var rows = await EnrichWithFlagImagesAsync(countries, log);
            await runningAppRepo.ExportToRunningAppAsync(hostApplication, rows, log);
        }

        public Task WriteDocumentJsonAsync(
            string exportType,
            string exportMethod,
            string sourcePath,
            string outputJsonPath,
            Action<string> log)
            => exportType switch
            {
                "Excel"      => _factory.GetExcelRepo(exportMethod).WriteDocumentJsonAsync(sourcePath, outputJsonPath, log),
                "PowerPoint" => _factory.GetPowerPointRepo(exportMethod).WriteDocumentJsonAsync(sourcePath, outputJsonPath, log),
                "Word"       => _factory.GetWordRepo(exportMethod).WriteDocumentJsonAsync(sourcePath, outputJsonPath, log),
                _            => Task.FromException(new NotSupportedException(
                                    $"Unsupported export type '{exportType}'."))
            };

        // ── Flag image enrichment ────────────────────────────────────────────────────

        private async Task<IList<MsOfficeCountry_Row>> EnrichWithFlagImagesAsync(
            IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            if (_flagImageManager == null || countries.All(c => c.FlagPng != null))
                return countries;

            log($"Resolving flag images for {countries.Count} countries (local cache first)...");
            var images = await _flagImageManager.GetFlagImagesAsync(
                countries.Where(c => c.FlagPng == null).Select(c => c.Iso2), log);

            return countries
                .Select(c => c.FlagPng == null && images.TryGetValue(c.Iso2, out var image)
                    ? c with { FlagPng = image.ImageBytes, FlagFilePath = image.FilePath }
                    : c)
                .ToList();
        }
    }
}
