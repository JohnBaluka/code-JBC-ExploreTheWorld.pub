using System;

namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Converts the Wikimedia SVG flag URLs returned by the CountriesNow.space API
    /// (<c>/countries/flag/images</c>) into rasterized PNG thumbnail URLs that can be
    /// embedded in Word/Excel/PowerPoint documents and downloaded for local caching.
    /// The API mixes Wikimedia Commons URLs (<c>/wikipedia/commons/…</c>) with English
    /// Wikipedia media URLs (<c>/wikipedia/en/…</c>); both use the same thumbnail scheme.
    /// </summary>
    public static class FlagImageUrl_Helper
    {
        /// <summary>
        /// Default thumbnail width in pixels. Wikimedia only serves a fixed list of
        /// thumbnail widths (see https://w.wiki/GHai) — 330 is on the list; arbitrary
        /// values such as 320 are rejected with HTTP 400.
        /// </summary>
        public const int DefaultWidthPx = 330;

        private const string WikipediaSegment = "/wikipedia/";

        /// <summary>
        /// Returns the Wikimedia PNG thumbnail URL for a Wikimedia SVG URL, e.g.
        /// <c>https://upload.wikimedia.org/wikipedia/commons/d/d4/Flag_of_Israel.svg</c> →
        /// <c>https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Flag_of_Israel.svg/330px-Flag_of_Israel.svg.png</c>.
        /// Returns the URL unchanged when it already points to a PNG/JPG bitmap, and
        /// <c>null</c> when the URL is empty or not a recognizable Wikimedia SVG URL.
        /// </summary>
        public static string? GetPngThumbnailUrl(string? flagUrl, int widthPx = DefaultWidthPx)
        {
            if (string.IsNullOrWhiteSpace(flagUrl) || widthPx <= 0)
                return null;

            var url = flagUrl!.Trim();

            if (url.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                url.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                url.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                return url;

            if (!url.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
                return null;

            // Split into ".../wikipedia/{project}/" prefix and the media-relative path.
            var segmentIndex = url.IndexOf(WikipediaSegment, StringComparison.OrdinalIgnoreCase);
            if (segmentIndex < 0)
                return null;

            var projectStart = segmentIndex + WikipediaSegment.Length;
            var projectEnd   = url.IndexOf('/', projectStart);
            if (projectEnd < 0)
                return null;

            var prefix       = url.Substring(0, projectEnd + 1);            // ".../wikipedia/{project}/"
            var relativePath = url.Substring(projectEnd + 1);               // "{h}/{hh}/{file}.svg" or "thumb/…"

            // Already a thumbnail path — append the sized PNG rendition of the SVG file name.
            if (relativePath.StartsWith("thumb/", StringComparison.OrdinalIgnoreCase))
                return $"{url}/{widthPx}px-{GetFileName(url)}.png";

            var fileName = GetFileName(relativePath);
            if (fileName.Length == 0)
                return null;

            return $"{prefix}thumb/{relativePath}/{widthPx}px-{fileName}.png";
        }

        private static string GetFileName(string url)
        {
            var slashIndex = url.LastIndexOf('/');
            return slashIndex >= 0 ? url.Substring(slashIndex + 1) : url;
        }
    }
}
