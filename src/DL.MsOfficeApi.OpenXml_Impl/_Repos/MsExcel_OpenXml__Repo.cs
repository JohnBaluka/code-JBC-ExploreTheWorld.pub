using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    public class MsExcel_OpenXml__Repo : MsExcel__Repo__Interface
    {
        // Data row height in points (fits the 0.25" flag image with a small margin).
        private const double FlagRowHeightPoints = 22;
        private const long FlagHeightEmu = 228600;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Creating Excel workbook (DocumentFormat.OpenXml)...");

                using var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new X.Workbook();
                log("  SpreadsheetDocument created.");

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData     = new X.SheetData();

                // Column A holds the flag images; make it wide enough for a 3:2 flag.
                var columns = new X.Columns(
                    new X.Column { Min = 1, Max = 1, Width = 7, CustomWidth = true });
                worksheetPart.Worksheet = new X.Worksheet(columns, sheetData);

                var sheets = workbookPart.Workbook.AppendChild(new X.Sheets());
                sheets.AppendChild(new X.Sheet
                {
                    Id      = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name    = "Countries"
                });
                log("  Worksheet 'Countries' created.");

                log("  Writing header row (4 columns)...");
                var headerRow = new X.Row { RowIndex = 1 };
                uint colIdx = 1;
                foreach (var col in new[] { "Flag", "Country", "ISO2", "ISO3" })
                {
                    headerRow.AppendChild(CreateXlCell(GetCellRef(colIdx, 1), col));
                    colIdx++;
                }
                sheetData.AppendChild(headerRow);

                log($"  Writing {countries.Count} data rows...");
                var worksheetDrawing = new Xdr.WorksheetDrawing();
                DrawingsPart? drawingsPart = null;
                uint rowIdx = 2;
                uint drawingId = 1;
                foreach (var c in countries)
                {
                    var dataRow = new X.Row { RowIndex = rowIdx };
                    if (c.FlagPng is { Length: > 0 })
                    {
                        dataRow.Height       = FlagRowHeightPoints;
                        dataRow.CustomHeight = true;

                        drawingsPart ??= worksheetPart.AddNewPart<DrawingsPart>();
                        var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
                        using (var stream = new MemoryStream(c.FlagPng))
                            imagePart.FeedData(stream);

                        var widthEmu = PngImage_Helper.GetWidthForHeight(c.FlagPng, FlagHeightEmu);
                        worksheetDrawing.AppendChild(OpenXmlFlagImage_Helper.CreateSpreadsheetAnchor(
                            drawingsPart.GetIdOfPart(imagePart), drawingId++, $"Flag {c.Iso2}",
                            columnIndex: 0, rowIndex: (int)rowIdx - 1, widthEmu, FlagHeightEmu));
                    }
                    dataRow.AppendChild(CreateXlCell(GetCellRef(2, rowIdx), c.Country));
                    dataRow.AppendChild(CreateXlCell(GetCellRef(3, rowIdx), c.Iso2));
                    dataRow.AppendChild(CreateXlCell(GetCellRef(4, rowIdx), c.Iso3));
                    sheetData.AppendChild(dataRow);
                    rowIdx++;
                    if ((rowIdx - 2) % 50 == 0 && rowIdx > 2)
                        log($"  Written {rowIdx - 2}/{countries.Count} rows...");
                }

                if (drawingsPart != null)
                {
                    drawingsPart.WorksheetDrawing = worksheetDrawing;
                    worksheetPart.Worksheet.AppendChild(new X.Drawing
                    {
                        Id = worksheetPart.GetIdOfPart(drawingsPart)
                    });
                    log($"  Embedded {drawingId - 1} flag images.");
                }

                log("  Saving workbook...");
                workbookPart.Workbook.Save();
                log($"  Saved: {filePath}");
            });
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
                JsonWriters.MsExcelJsonWriter.WriteWorkbookFileToJson(sourcePath, outputJsonPath, null, log));
        }

        // ── Cell builder helpers ─────────────────────────────────────────────

        private static X.Cell CreateXlCell(string reference, string value)
        {
            var cell = new X.Cell { CellReference = reference, DataType = X.CellValues.InlineString };
            cell.AppendChild(new X.InlineString(new X.Text(value)));
            return cell;
        }

        private static string GetCellRef(uint col, uint row)
        {
            var name = string.Empty;
            var c = col;
            while (c > 0)
            {
                var mod = (c - 1) % 26;
                name = (char)('A' + mod) + name;
                c    = (c - 1) / 26;
            }
            return $"{name}{row}";
        }
    }
}
