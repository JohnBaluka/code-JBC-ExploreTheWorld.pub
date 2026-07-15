using System;
using System.IO;
using System.Linq;

namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Maps a OneDrive cloud URL (the path Office reports in <c>Document.FullName</c> for
    /// AutoSave documents) to the corresponding file in the local OneDrive sync folder, e.g.
    /// <c>https://d.docs.live.net/{cid}/Presentations/Deck.pptx</c> →
    /// <c>C:\Users\{user}\OneDrive\Presentations\Deck.pptx</c>.
    /// The sync roots come from the %OneDriveConsumer% / %OneDriveCommercial% / %OneDrive%
    /// environment variables that the OneDrive sync client maintains.
    /// </summary>
    public static class OneDriveLocalPath_Helper
    {
        public static string? TryGetLocalPath(string? cloudUrl)
        {
            return TryGetLocalPath(cloudUrl, Environment.GetEnvironmentVariable);
        }

        // Overload with an environment-variable resolver so tests can supply their own roots.
        public static string? TryGetLocalPath(string? cloudUrl, Func<string, string?> getEnvironmentVariable)
        {
            if (string.IsNullOrWhiteSpace(cloudUrl)) return null;
            if (!cloudUrl!.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return null;
            if (!Uri.TryCreate(cloudUrl, UriKind.Absolute, out var uri)) return null;

            string relativePath;
            string[] rootVariables;

            if (uri.Host.Equals("d.docs.live.net", StringComparison.OrdinalIgnoreCase))
            {
                // Personal OneDrive: https://d.docs.live.net/{cid}/{path...}
                relativePath = RelativePathAfterSegments(uri, 1);
                rootVariables = new[] { "OneDriveConsumer", "OneDrive" };
            }
            else if (uri.Host.EndsWith("-my.sharepoint.com", StringComparison.OrdinalIgnoreCase))
            {
                // OneDrive for Business: https://{tenant}-my.sharepoint.com/personal/{user}/Documents/{path...}
                relativePath = RelativePathAfterSegments(uri, 3);
                rootVariables = new[] { "OneDriveCommercial", "OneDrive" };
            }
            else
            {
                return null;
            }

            if (relativePath.Length == 0) return null;

            foreach (var variable in rootVariables)
            {
                var root = getEnvironmentVariable(variable);
                if (string.IsNullOrEmpty(root)) continue;

                var candidate = Path.Combine(root, relativePath);
                if (File.Exists(candidate) || Directory.Exists(candidate)) return candidate;
            }

            return null;
        }

        // Returns the URL path after skipping the first N segments, URL-decoded and using backslashes.
        private static string RelativePathAfterSegments(Uri uri, int skipCount)
        {
            var segments = uri.AbsolutePath
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Skip(skipCount)
                .Select(Uri.UnescapeDataString);

            return string.Join(@"\", segments);
        }
    }
}
