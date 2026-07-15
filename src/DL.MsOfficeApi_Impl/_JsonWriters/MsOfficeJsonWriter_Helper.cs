using System;
using DynamicWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters;
using NetOfficeWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    /// <summary>
    /// Writes the structure of an open MS Office document/workbook/presentation to a JSON file.
    /// All methods accept <c>dynamic</c> so they work with .NET Interop PIAs, Late-Binding, and
    /// NetOffice objects: NetOffice objects go to the strongly-typed DL.MsOfficeApi.NetOffice_Impl writers,
    /// anything else to the late-bound DL.MsOfficeApi.Dynamic_Impl writers. Both emit the canonical
    /// JBC.ExploreTheWorld.DL.Ms{Word|Excel|PowerPoint} entity schema (same as VBA).
    /// </summary>
    public static class MsOfficeJsonWriter_Helper
    {
        public static void WriteWordJson(dynamic doc, string filePath, Action<string> log)
        {
            if (doc is NetOffice.WordApi.Document netOfficeDocument)
            {
                NetOfficeWriters.MsWordJsonWriter.WriteDocumentToJsonFile(netOfficeDocument, filePath, null, log);
                return;
            }

            DynamicWriters.MsWordJsonWriter.WriteDocumentComToJsonFile(doc, filePath, null, log);
        }

        public static void WriteExcelJson(dynamic wb, string filePath, Action<string> log)
        {
            if (wb is NetOffice.ExcelApi.Workbook netOfficeWorkbook)
            {
                NetOfficeWriters.MsExcelJsonWriter.WriteWorkbookToJsonFile(netOfficeWorkbook, filePath, null, log);
                return;
            }

            DynamicWriters.MsExcelJsonWriter.WriteWorkbookComToJsonFile(wb, filePath, null, log);
        }

        public static void WritePowerPointJson(dynamic prs, string filePath, Action<string> log)
        {
            if (prs is NetOffice.PowerPointApi.Presentation netOfficePresentation)
            {
                NetOfficeWriters.MsPowerPointJsonWriter.WritePresentationToJsonFile(netOfficePresentation, filePath, null, log);
                return;
            }

            DynamicWriters.MsPowerPointJsonWriter.WritePresentationComToJsonFile(prs, filePath, null, log);
        }
    }
}
