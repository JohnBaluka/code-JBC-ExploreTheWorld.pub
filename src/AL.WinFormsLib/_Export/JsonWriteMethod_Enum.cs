namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    // Order mirrors the MsOfficeApi picklist order (COM, Direct, Interop, Dynamic, NetOffice,
    // OpenXML) for the members that apply to Save-As-JSON.
    public enum JsonWriteMethod
    {
        Direct,
        Interop,
        Dynamic,
        NetOffice,
        OpenXml
    }
}
