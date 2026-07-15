using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using NetOffice.ExcelApi;
using ERange = NetOffice.ExcelApi.Range;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl
{
    public class MsExcel_NetOffice__Repo : MsExcel__Repo__Interface, MsOfficeRunningAppExport__Repo__Interface
    {
        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Launching Microsoft Excel via NetOffice...");

                using var excelApp = new Application { Visible = true };
                excelApp.DisplayAlerts = false;

                excelApp.WorkbookBeforeCloseEvent += (Workbook wb, ref bool cancel) =>
                    log($"  [Event] WorkbookBeforeClose  — {wb.Name}");
                excelApp.WorkbookBeforeSaveEvent += (Workbook wb, bool saveAsUi, ref bool cancel) =>
                    log($"  [Event] WorkbookBeforeSave   — {wb.Name}");
                excelApp.WorkbookOpenEvent += wb =>
                    log($"  [Event] WorkbookOpen         — {wb.Name}");

                log("  Creating new workbook...");
                var wb2 = excelApp.Workbooks.Add();
                log($"  [Event] New workbook created: {wb2.Name}");

                FillWorkbook(wb2, countries, log);

                log($"  Saving: {filePath}");
                wb2.SaveAs(filePath, NetOffice.ExcelApi.Enums.XlFileFormat.xlOpenXMLWorkbook);
                log("  Closing Excel...");
                excelApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        // Creates a new workbook in the caller's already-running Excel application and fills it,
        // leaving it open and unsaved. Runs synchronously on the caller's (STA/UI) thread — the
        // host COM object is bound to that apartment and must not be marshalled onto a pool thread.
        public Task ExportToRunningAppAsync(object hostApplication, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            if (hostApplication is not Application excelApp)
                throw new ArgumentException(
                    "Expected a NetOffice.ExcelApi.Application instance.", nameof(hostApplication));

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
                var flagPath = NetOfficeFlagImage_Helper.GetFlagImageFilePath(c);
                if (flagPath != null)
                {
                    var cell = (ERange)ws.Cells[r + 1, 1];
                    cell.RowHeight = flagRowHeightPoints;
                    // Insert at original size, then scale by height with the aspect ratio locked.
                    var picture = ws.Shapes.AddPicture(flagPath,
                        NetOffice.OfficeApi.Enums.MsoTriState.msoFalse,
                        NetOffice.OfficeApi.Enums.MsoTriState.msoCTrue,
                        Convert.ToSingle(cell.Left) + 2f,
                        Convert.ToSingle(cell.Top) + 2f,
                        -1f, -1f);
                    picture.LockAspectRatio = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
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
                log($"Writing (NetOffice): {outputJsonPath}");

                using var excelApp = new Application { Visible = true };
                excelApp.DisplayAlerts = false;
                var wb = excelApp.Workbooks.Open(sourcePath);
                try
                {
                    WriteJsonFromOpenWorkbook(wb, outputJsonPath, log);
                }
                finally
                {
                    wb.Close(false);
                    excelApp.Quit();
                }
            });
        }

        public void WriteJsonFromOpenWorkbook(Workbook wb, string outputJsonPath, Action<string> log)
            => JsonWriters.MsExcelJsonWriter.WriteWorkbookToJsonFile(wb, outputJsonPath, null, log);
    }
}