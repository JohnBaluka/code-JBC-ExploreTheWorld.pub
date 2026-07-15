namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    /// <summary>
    /// Supplies the concrete Office document repository for a given export library
    /// ("OpenXML", "NetOffice", "Interop"). A DL contract implemented by the host application
    /// layer (the only layer allowed to reference the platform-specific DL repo `_Impl`
    /// projects), so the BL export service — and the UI that calls it — never instantiate
    /// those repos directly.
    /// </summary>
    public interface MsOfficeExportRepoFactory__Interface
    {
        MsWord__Repo__Interface GetWordRepo(string exportMethod);
        MsExcel__Repo__Interface GetExcelRepo(string exportMethod);
        MsPowerPoint__Repo__Interface GetPowerPointRepo(string exportMethod);
    }
}
