using System;
using System.IO;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl
{
    /// <summary>
    /// COM <c>AddPicture</c> calls require a file on disk. This helper returns the local
    /// cache path carried on the row when available, and otherwise materializes the PNG
    /// bytes to a per-user temp folder (<c>%TEMP%\ETW_FlagImages</c>) that acts as a
    /// secondary cache for COM-based exports.
    /// </summary>
    internal static class NetOfficeFlagImage_Helper
    {
        public static string? GetFlagImageFilePath(MsOfficeCountry_Row country)
        {
            if (!string.IsNullOrEmpty(country.FlagFilePath) && File.Exists(country.FlagFilePath))
                return country.FlagFilePath;

            if (country.FlagPng is not { Length: > 0 } || string.IsNullOrWhiteSpace(country.Iso2))
                return null;

            var folder = Path.Combine(Path.GetTempPath(), "ETW_FlagImages");
            Directory.CreateDirectory(folder);

            var path = Path.Combine(folder, country.Iso2.Trim().ToUpperInvariant() + ".png");
            if (!File.Exists(path) || new FileInfo(path).Length != country.FlagPng.Length)
                File.WriteAllBytes(path, country.FlagPng);
            return path;
        }
    }
}
