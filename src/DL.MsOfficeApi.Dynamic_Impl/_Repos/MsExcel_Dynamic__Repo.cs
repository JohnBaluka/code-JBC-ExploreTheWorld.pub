using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl
{
    public class MsExcel_Dynamic__Repo : MsExcel__Repo__Interface
    {
        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Connecting to Microsoft Excel via dynamic COM automation...");

                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Excel not found. Ensure Microsoft Office is installed.");

                log("  Creating Excel application instance...");
                dynamic excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible       = true; // keep the app visible while exporting
                excelApp.DisplayAlerts = false;
                log($"  [Event] Connected — Excel {excelApp.Version}");

                log("  Creating new workbook...");
                dynamic wb = excelApp.Workbooks.Add();
                log($"  [Event] WorkbookOpen — {wb.Name}");

                dynamic ws = wb.Worksheets[1];
                ws.Name = "Countries";

                const float flagHeightPoints = 16f;
                const double flagRowHeightPoints = 20;

                log($"  Writing header row (4 columns)...");
                ws.Cells[1, 1] = "Flag";
                ws.Cells[1, 2] = "Country";
                ws.Cells[1, 3] = "ISO2";
                ws.Cells[1, 4] = "ISO3";
                ws.Columns[1].ColumnWidth = 7;

                log($"  Writing {countries.Count} data rows...");
                int r = 0;
                int flagCount = 0;
                foreach (var c in countries)
                {
                    r++;
                    var flagPath = DynamicFlagImage_Helper.GetFlagImageFilePath(c);
                    if (flagPath != null)
                    {
                        dynamic cell = ws.Cells[r + 1, 1];
                        cell.RowHeight = flagRowHeightPoints;
                        // Insert at original size, then scale by height with the aspect ratio locked.
                        dynamic picture = ws.Shapes.AddPicture(flagPath,
                            0,   // msoFalse   — do not link to file
                            1,   // msoCTrue   — save image with the document
                            Convert.ToSingle(cell.Left) + 2f,
                            Convert.ToSingle(cell.Top) + 2f,
                            -1f, -1f);
                        picture.LockAspectRatio = -1; // msoTrue
                        picture.Height = flagHeightPoints;
                        flagCount++;
                    }
                    ws.Cells[r + 1, 2] = c.Country;
                    ws.Cells[r + 1, 3] = c.Iso2;
                    ws.Cells[r + 1, 4] = c.Iso3;
                    if (r % 50 == 0)
                        log($"  Written {r}/{countries.Count} rows...");
                }
                log($"  Embedded {flagCount} flag images.");

                log($"  [Event] WorkbookBeforeSave — {wb.Name}");
                log($"  Saving: {filePath}");
                wb.SaveAs(filePath, 51); // xlOpenXMLWorkbook = 51
                log("  [Event] Workbook saved.");
                wb.Close(false);
                excelApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (Dynamic): {outputJsonPath}");

                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Excel not found. Ensure Microsoft Office is installed.");

                dynamic excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible       = true;
                excelApp.DisplayAlerts = false;
                try
                {
                    dynamic wb = excelApp.Workbooks.Open(sourcePath);
                    try
                    {
                        JsonWriters.MsExcelJsonWriter.WriteWorkbookComToJsonFile(wb, outputJsonPath, null, log);
                    }
                    finally
                    {
                        wb.Close(false);
                    }
                }
                finally
                {
                    excelApp.Quit();
                }
            });
        }
    }
}
