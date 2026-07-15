using System;
using System.IO;
using NetOffice.ExcelApi;
using DynamicWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters;
using InteropWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl.JsonWriters;
using NetOfficeWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters;
using OpenXmlWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    // Thin wrapper over the DL "Save as JSON" writers so the watcher forms keep their
    // existing call sites. Both paths emit the canonical
    // JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entity schema (same as the VBA writers).
    internal static class MsExcel_JsonWriter
    {
        /// <summary>Writes a full JSON representation of the workbook using the NetOffice COM object model.</summary>
        internal static void WriteNetOffice(Workbook wb, string filePath, Action<string> log)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            NetOfficeWriters.MsExcelJsonWriter.WriteWorkbookToJsonFile(wb, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the workbook using the late-bound dynamic COM writers (DL.MsOfficeApi.Dynamic_Impl).</summary>
        internal static void WriteDynamic(Workbook wb, string filePath, Action<string> log)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            DynamicWriters.MsExcelJsonWriter.WriteWorkbookComToJsonFile(wb.UnderlyingObject, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the workbook using the strongly-typed Microsoft.Office.Interop COM object model (DL.MsOfficeApi.Interop_Impl).</summary>
        internal static void WriteInterop(Workbook wb, string filePath, Action<string> log)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            // The watcher holds a NetOffice wrapper; unwrap to the raw COM object and cast to the
            // Interop interface (a QueryInterface on the same live Excel.Workbook COM object).
            var interopWorkbook = (Microsoft.Office.Interop.Excel.Workbook)wb.UnderlyingObject;
            InteropWriters.MsExcelJsonWriter.WriteWorkbookToJsonFile(interopWorkbook, filePath, null, log);
        }

        /// <summary>Writes a JSON representation of the .xlsx file at <paramref name="sourceXlsxPath"/> using OpenXML (no COM required).</summary>
        internal static void WriteOpenXml(string sourceXlsxPath, string outputJsonFilePath, Action<string> log)
        {
            if (string.IsNullOrEmpty(sourceXlsxPath)) throw new ArgumentException("sourceXlsxPath is required.", nameof(sourceXlsxPath));
            if (string.IsNullOrEmpty(outputJsonFilePath)) throw new ArgumentException("outputJsonFilePath is required.", nameof(outputJsonFilePath));

            OpenXmlWriters.MsExcelJsonWriter.WriteWorkbookFileToJson(sourceXlsxPath, outputJsonFilePath, null, log);
        }

        /// <summary>
        /// Writes the running app's active workbook to JSON using OpenXML. OpenXML reads the file
        /// on disk while Excel holds a lock on it, so the workbook is saved, closed, read, and
        /// reopened. Cloud (AutoSave/OneDrive) workbooks have an https path OpenXML cannot read;
        /// those are read from a temporary local copy instead (no close/reopen).
        /// </summary>
        internal static void WriteOpenXmlFromRunningApp(NetOffice.ExcelApi.Application excelApp, string outputJsonFilePath, Action<string> log)
        {
            if (excelApp == null) throw new ArgumentNullException(nameof(excelApp));

            var wb = excelApp.ActiveWorkbook
                ?? throw new InvalidOperationException("No active workbook.");

            string documentFolder = string.Empty;
            try { documentFolder = wb.Path; } catch { }
            if (string.IsNullOrEmpty(documentFolder))
                throw new InvalidOperationException(
                    "The active workbook has never been saved — save it first, or use the NetOffice method.");

            string fullName = wb.FullName;

            if (fullName.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                log("Cloud (AutoSave) workbook detected — reading a temporary local copy...");
                string extension = Path.GetExtension(fullName);
                if (string.IsNullOrEmpty(extension)) extension = ".xlsx";
                string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + extension);

                wb.SaveCopyAs(tempPath);
                try
                {
                    WriteOpenXml(tempPath, outputJsonFilePath, log);
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }
                return;
            }

            log("Closing active workbook for OpenXML read...");
            try { wb.Save(); } catch { }
            wb.Close(false);
            try
            {
                WriteOpenXml(fullName, outputJsonFilePath, log);
            }
            finally
            {
                log("Reopening workbook...");
                excelApp.Workbooks.Open(fullName);
            }
        }
    }
}
