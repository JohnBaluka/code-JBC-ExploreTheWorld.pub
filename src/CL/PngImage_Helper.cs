namespace JBC.ExploreTheWorld.CL
{
    /// <summary>
    /// Minimal PNG header reader used to size embedded flag images without a graphics
    /// dependency (works on .NET 10, .NET Framework 4.8.1, and WASM).
    /// </summary>
    public static class PngImage_Helper
    {
        /// <summary>
        /// Reads the pixel width/height from a PNG byte stream (IHDR chunk).
        /// Returns <c>false</c> when the bytes are not a valid PNG header.
        /// </summary>
        public static bool TryGetPixelSize(byte[]? pngBytes, out int width, out int height)
        {
            width  = 0;
            height = 0;

            // 8-byte signature + 4 length + 4 "IHDR" + 4 width + 4 height = 24 bytes minimum.
            if (pngBytes == null || pngBytes.Length < 24)
                return false;

            if (pngBytes[0] != 0x89 || pngBytes[1] != 0x50 || pngBytes[2] != 0x4E || pngBytes[3] != 0x47 ||
                pngBytes[12] != (byte)'I' || pngBytes[13] != (byte)'H' || pngBytes[14] != (byte)'D' || pngBytes[15] != (byte)'R')
                return false;

            width  = (pngBytes[16] << 24) | (pngBytes[17] << 16) | (pngBytes[18] << 8) | pngBytes[19];
            height = (pngBytes[20] << 24) | (pngBytes[21] << 16) | (pngBytes[22] << 8) | pngBytes[23];
            return width > 0 && height > 0;
        }

        /// <summary>
        /// Returns the display width for the given display height, preserving the PNG's
        /// aspect ratio; falls back to a 3:2 flag ratio when the size cannot be read.
        /// </summary>
        public static long GetWidthForHeight(byte[]? pngBytes, long displayHeight)
        {
            if (TryGetPixelSize(pngBytes, out var width, out var height))
                return (long)((double)displayHeight * width / height);
            return displayHeight * 3 / 2;
        }
    }
}
