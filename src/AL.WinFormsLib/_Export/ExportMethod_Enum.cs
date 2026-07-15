namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    // Order mirrors the MsOfficeApi picklist order for the members that apply to Export Data
    // (Direct is Save-As-JSON only, so it is not an export method).
    public enum ExportMethod
    {
        Interop,
        Dynamic,
        NetOffice,
        OpenXml
    }
}
