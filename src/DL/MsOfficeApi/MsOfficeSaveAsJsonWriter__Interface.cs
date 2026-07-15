using System;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    /// <summary>
    /// Writes the structure of a live (COM-connected) Office document/workbook/presentation to a
    /// JSON file, dispatching on the user-selected write method ("NetOffice", "Interop", "Dynamic",
    /// "OpenXml", "Direct"). A DL contract implemented by the host application layer (the only layer
    /// allowed to reference the platform-specific DL repo `_Impl` projects), so the AL WinForms UI
    /// library — the watcher forms that call it — never references those `_Impl` projects directly.
    /// <para>
    /// The COM objects are passed as <see cref="object"/> (the core DL cannot reference NetOffice or
    /// the Office PIAs); the implementation casts them to the concrete COM types. Callers pass the
    /// live document/workbook/presentation (used by NetOffice/Interop/Dynamic) and the owning
    /// application (used by OpenXml/Direct); either may be <c>null</c> when not applicable.
    /// </para>
    /// </summary>
    public interface MsOfficeSaveAsJsonWriter__Interface
    {
        void WriteWordJson(string method, object? document, object? application, string outputFilePath, Action<string> log);
        void WriteExcelJson(string method, object? workbook, object? application, string outputFilePath, Action<string> log);
        void WritePowerPointJson(string method, object? presentation, object? application, string outputFilePath, Action<string> log);
    }
}
