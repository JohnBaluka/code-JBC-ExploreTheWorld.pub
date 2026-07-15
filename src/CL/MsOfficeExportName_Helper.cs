namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Builds default "Export Data" file names that embed the Office API method (and, when the
    /// source country data came from the Access database, an "Access" token) so the output name
    /// records how it was produced — e.g.
    /// <c>("ETW_CountriesNow", "Direct", fromAccessDb: true, "pptx")</c> →
    /// <c>ETW_CountriesNow-Access-Direct.pptx</c>, and without Access →
    /// <c>ETW_CountriesNow-Direct.pptx</c>.
    /// </summary>
    public static class MsOfficeExportName_Helper
    {
        public static string BuildFileName(string baseName, string method, bool fromAccessDb, string extension)
            => $"{BuildStem(baseName, method, fromAccessDb)}.{extension}";

        public static string BuildStem(string baseName, string method, bool fromAccessDb)
        {
            var accessToken = fromAccessDb ? "-Access" : string.Empty;
            var methodToken = string.IsNullOrWhiteSpace(method) ? string.Empty : $"-{method}";
            return $"{baseName}{accessToken}{methodToken}";
        }
    }
}
