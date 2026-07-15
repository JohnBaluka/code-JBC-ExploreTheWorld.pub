using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using ERange = Microsoft.Office.Interop.Excel.Range;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl
{
    public class MsExcel_Interop__Repo : MsExcel__Repo__Interface, MsOfficeRunningAppExport__Repo__Interface
    {
        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Launching Microsoft Excel via Interop...");

                var excelApp = new Application { Visible = true };
                excelApp.DisplayAlerts = false;
                bool quit = false;
                try
                {
                    excelApp.WorkbookBeforeClose += (Workbook wb, ref bool cancel) =>
                        log($"  [Event] WorkbookBeforeClose  — {wb.Name}");
                    excelApp.WorkbookBeforeSave += (Workbook wb, bool saveAsUi, ref bool cancel) =>
                        log($"  [Event] WorkbookBeforeSave   — {wb.Name}");
                    excelApp.WorkbookOpen += wb =>
                        log($"  [Event] WorkbookOpen         — {wb.Name}");

                    log("  Creating new workbook...");
                    var wb2 = excelApp.Workbooks.Add();
                    log($"  [Event] New workbook created: {wb2.Name}");

                    FillWorkbook(wb2, countries, log);

                    log($"  Saving: {filePath}");
                    wb2.SaveAs(filePath, XlFileFormat.xlOpenXMLWorkbook);
                    log("  Closing Excel...");
                    excelApp.Quit();
                    quit = true;
                    log($"  Saved: {filePath}");
                }
                finally
                {
                    if (!quit) { try { excelApp.Quit(); } catch { /* best effort */ } }
                    Marshal.FinalReleaseComObject(excelApp);
                }
            });
        }

        // Creates a new workbook in the caller's already-running Excel application and fills it,
        // leaving it open and unsaved. Runs synchronously on the caller's (STA/UI) thread — the
        // host COM object is bound to that apartment and must not be marshalled onto a pool thread.
        public Task ExportToRunningAppAsync(object hostApplication, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            if (hostApplication is not Application excelApp)
                throw new ArgumentException(
                    "Expected a Microsoft.Office.Interop.Excel.Application instance.", nameof(hostApplication));

            log("Creating a new workbook in the running Excel application...");
            var wb = excelApp.Workbooks.Add();
            FillWorkbook(wb, countries, log);
            log($"  New workbook created: {wb.Name}");
            return Task.CompletedTask;
        }

        private static void FillWorkbook(Workbook wb, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            const float flagHeightPoints = 16f;
            const double flagRowHeightPoints = 20;

            var ws = (Worksheet)wb.Worksheets[1];
            ws.Name = "Countries";

            log("  Writing header row (4 columns)...");
            ((ERange)ws.Cells[1, 1]).Value2 = "Flag";
            ((ERange)ws.Cells[1, 2]).Value2 = "Country";
            ((ERange)ws.Cells[1, 3]).Value2 = "ISO2";
            ((ERange)ws.Cells[1, 4]).Value2 = "ISO3";
            ((ERange)ws.Columns[1]).ColumnWidth = 7;

            log($"  Writing {countries.Count} data rows...");
            int r = 0;
            int flagCount = 0;
            foreach (var c in countries)
            {
                r++;
                var flagPath = InteropFlagImage_Helper.GetFlagImageFilePath(c);
                if (flagPath != null)
                {
                    var cell = (ERange)ws.Cells[r + 1, 1];
                    cell.RowHeight = flagRowHeightPoints;
                    // Insert at original size, then scale by height with the aspect ratio locked.
                    var picture = ws.Shapes.AddPicture(flagPath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoCTrue,
                        Convert.ToSingle(cell.Left) + 2f,
                        Convert.ToSingle(cell.Top) + 2f,
                        -1f, -1f);
                    picture.LockAspectRatio = MsoTriState.msoTrue;
                    picture.Height = flagHeightPoints;
                    flagCount++;
                }
                ((ERange)ws.Cells[r + 1, 2]).Value2 = c.Country;
                ((ERange)ws.Cells[r + 1, 3]).Value2 = c.Iso2;
                ((ERange)ws.Cells[r + 1, 4]).Value2 = c.Iso3;
                if (r % 50 == 0)
                    log($"  Written {r}/{countries.Count} rows...");
            }
            log($"  Embedded {flagCount} flag images.");
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (Interop): {outputJsonPath}");

                var excelApp = new Application { Visible = true };
                excelApp.DisplayAlerts = false;
                try
                {
                    var wb = excelApp.Workbooks.Open(sourcePath);
                    try
                    {
                        WriteJsonFromOpenWorkbook(wb, outputJsonPath, log);
                    }
                    finally
                    {
                        wb.Close(false);
                    }
                }
                finally
                {
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);
                }
            });
        }

        public void WriteJsonFromOpenWorkbook(Workbook wb, string outputJsonPath, Action<string> log)
            => JsonWriters.MsExcelJsonWriter.WriteWorkbookToJsonFile(wb, outputJsonPath, null, log);
    }
}
