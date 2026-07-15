using System;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    /// <summary>
    /// One country row passed to the Word/Excel/PowerPoint export writers.
    /// <paramref name="FlagPng"/> holds the PNG flag image bytes to embed (or <c>null</c>
    /// for a text-only row); <paramref name="FlagFilePath"/> is the local cache path of the
    /// same image when available — COM-based writers (NetOffice/VBA) prefer the path.
    /// </summary>
    public record MsOfficeCountry_Row(
        string Country,
        string Iso2,
        string Iso3,
        byte[]? FlagPng = null,
        string? FlagFilePath = null);
}
