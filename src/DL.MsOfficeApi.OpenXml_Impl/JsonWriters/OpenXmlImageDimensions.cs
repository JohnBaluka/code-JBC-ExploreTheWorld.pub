using System;
using System.IO;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Minimal header parsing for the raster formats PowerPoint embeds, returning the
    // pixel size and DPI needed to convert srcRect crop percentages into points.
    internal sealed class OpenXmlImageDimensions
    {
        public int PixelWidth { get; private set; }
        public int PixelHeight { get; private set; }
        public double DpiX { get; private set; } = 96.0;
        public double DpiY { get; private set; } = 96.0;

        public double WidthPoints
        {
            get
            {
                return PixelWidth * 72.0 / DpiX;
            }
        }

        public double HeightPoints
        {
            get
            {
                return PixelHeight * 72.0 / DpiY;
            }
        }

        public static OpenXmlImageDimensions? Parse(byte[] bytes)
        {
            try
            {
                if (bytes.Length > 24 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                {
                    return ParsePng(bytes);
                }

                if (bytes.Length > 4 && bytes[0] == 0xFF && bytes[1] == 0xD8)
                {
                    return ParseJpeg(bytes);
                }

                if (bytes.Length > 10 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F')
                {
                    return new OpenXmlImageDimensions
                    {
                        PixelWidth = bytes[6] | (bytes[7] << 8),
                        PixelHeight = bytes[8] | (bytes[9] << 8),
                    };
                }

                if (bytes.Length > 40 && bytes[0] == 'B' && bytes[1] == 'M')
                {
                    return ParseBmp(bytes);
                }
            }
            catch
            {
                // Unparseable headers fall through to null.
            }

            return null;
        }

        private static OpenXmlImageDimensions ParsePng(byte[] bytes)
        {
            static int ReadInt32BigEndian(byte[] data, int offset)
            {
                return (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
            }

            var result = new OpenXmlImageDimensions
            {
                PixelWidth = ReadInt32BigEndian(bytes, 16),
                PixelHeight = ReadInt32BigEndian(bytes, 20),
            };

            // Walk the chunks for pHYs (pixels per meter).
            int position = 8;
            while (position + 8 < bytes.Length)
            {
                int length = ReadInt32BigEndian(bytes, position);
                string type = System.Text.Encoding.ASCII.GetString(bytes, position + 4, 4);

                if (type == "pHYs" && position + 8 + 9 <= bytes.Length)
                {
                    int perMeterX = ReadInt32BigEndian(bytes, position + 8);
                    int perMeterY = ReadInt32BigEndian(bytes, position + 12);
                    byte unit = bytes[position + 16];

                    if (unit == 1 && perMeterX > 0 && perMeterY > 0)
                    {
                        result.DpiX = perMeterX * 0.0254;
                        result.DpiY = perMeterY * 0.0254;
                    }

                    break;
                }

                if (type == "IDAT" || type == "IEND") break;

                position += 12 + length;
                if (length < 0) break;
            }

            return result;
        }

        private static OpenXmlImageDimensions? ParseJpeg(byte[] bytes)
        {
            var result = new OpenXmlImageDimensions();
            bool sizeFound = false;

            int position = 2;
            while (position + 4 < bytes.Length)
            {
                if (bytes[position] != 0xFF)
                {
                    position++;
                    continue;
                }

                byte marker = bytes[position + 1];
                if (marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD9))
                {
                    position += 2;
                    continue;
                }

                int length = (bytes[position + 2] << 8) | bytes[position + 3];

                if (marker == 0xE0 && length >= 14 && position + 2 + length <= bytes.Length)
                {
                    // JFIF: units, Xdensity, Ydensity.
                    byte units = bytes[position + 11];
                    int densityX = (bytes[position + 12] << 8) | bytes[position + 13];
                    int densityY = (bytes[position + 14] << 8) | bytes[position + 15];

                    if (densityX > 0 && densityY > 0)
                    {
                        if (units == 1)
                        {
                            result.DpiX = densityX;
                            result.DpiY = densityY;
                        }
                        else if (units == 2)
                        {
                            result.DpiX = densityX * 2.54;
                            result.DpiY = densityY * 2.54;
                        }
                    }
                }

                bool isSofMarker = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                if (isSofMarker && position + 9 <= bytes.Length)
                {
                    result.PixelHeight = (bytes[position + 5] << 8) | bytes[position + 6];
                    result.PixelWidth = (bytes[position + 7] << 8) | bytes[position + 8];
                    sizeFound = true;
                    break;
                }

                position += 2 + length;
            }

            return sizeFound ? result : null;
        }

        private static OpenXmlImageDimensions ParseBmp(byte[] bytes)
        {
            var result = new OpenXmlImageDimensions
            {
                PixelWidth = BitConverter.ToInt32(bytes, 18),
                PixelHeight = Math.Abs(BitConverter.ToInt32(bytes, 22)),
            };

            int perMeterX = BitConverter.ToInt32(bytes, 38);
            int perMeterY = BitConverter.ToInt32(bytes, 42);
            if (perMeterX > 0 && perMeterY > 0)
            {
                result.DpiX = perMeterX * 0.0254;
                result.DpiY = perMeterY * 0.0254;
            }

            return result;
        }
    }
}
