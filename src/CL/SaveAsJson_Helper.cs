using System;
using System.IO;
using System.Linq;

namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Builds the default "Save As JSON" output path. The default file name is the current
    /// document's file name with the Save-As-JSON <paramref name="method"/> inserted before the
    /// extension and ".json" appended (e.g. <c>ETW_CountriesNow.docx</c> + method <c>NetOffice</c>
    /// → <c>ETW_CountriesNow-NetOffice.docx.json</c>), placed in the current document's folder when
    /// it is known, otherwise the user's Documents folder. When <paramref name="method"/> is null
    /// the method token is omitted (legacy <c>{name}.json</c> form). Cloud (AutoSave/OneDrive)
    /// documents report an https path; that is mapped to the local OneDrive sync folder when
    /// possible, otherwise the default falls back to the Documents folder.
    /// </summary>
    public static class SaveAsJson_Helper
    {
        public static string BuildDefaultPath(string? currentFilePath, string fallbackFileName, string? method = null)
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (string.IsNullOrWhiteSpace(currentFilePath))
                return Path.Combine(documents, BuildJsonName(fallbackFileName, method));

            if (currentFilePath!.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                var localPath = OneDriveLocalPath_Helper.TryGetLocalPath(currentFilePath);
                if (localPath != null)
                {
                    currentFilePath = localPath;
                }
                else
                {
                    // Unknown cloud location — default to Documents with the document's file name.
                    var cloudFileName = Uri.UnescapeDataString(
                        currentFilePath.Split('/').LastOrDefault(s => s.Length > 0) ?? string.Empty);
                    if (string.IsNullOrEmpty(cloudFileName)) cloudFileName = fallbackFileName;

                    return Path.Combine(documents, BuildJsonName(cloudFileName, method));
                }
            }

            var fileName  = Path.GetFileName(currentFilePath);
            var directory = Path.GetDirectoryName(currentFilePath);

            // A bare document name (e.g. an unsaved "Document1") has no directory — fall back to Documents.
            if (string.IsNullOrEmpty(directory))
                directory = documents;

            return Path.Combine(directory, BuildJsonName(fileName, method));
        }

        // Inserts the "-{method}" token before the document extension, then appends ".json"
        // (e.g. "Report.pptx" + "Direct" → "Report-Direct.pptx.json"). With no method the legacy
        // "{name}.json" form is preserved.
        private static string BuildJsonName(string fileName, string? method)
        {
            if (string.IsNullOrWhiteSpace(method))
                return fileName + ".json";

            var stem = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            return $"{stem}-{method}{extension}.json";
        }
    }
}
