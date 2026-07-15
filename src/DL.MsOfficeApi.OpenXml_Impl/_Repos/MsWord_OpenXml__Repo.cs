using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    public class MsWord_OpenXml__Repo : MsWord__Repo__Interface
    {
        // Flag display height in the table cell (0.25 inch; width follows the PNG aspect ratio).
        private const long FlagHeightEmu = 228600;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Creating Word document (DocumentFormat.OpenXml)...");

                using var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
                log("  WordprocessingDocument created.");

                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new W.Document();
                var body = mainPart.Document.AppendChild(new W.Body());
                log("  MainDocumentPart added.");

                log("  Adding title paragraph...");
                var titleRun = new W.Run(new W.Text("Countries Now"));
                titleRun.RunProperties = new W.RunProperties(new W.Bold(), new W.FontSize { Val = "36" });
                body.AppendChild(new W.Paragraph(titleRun));
                body.AppendChild(new W.Paragraph());

                log($"  Building table (4 columns, {countries.Count} rows)...");
                var table = new W.Table();
                table.AppendChild(new W.TableProperties(
                    new W.TableBorders(
                        new W.TopBorder              { Val = W.BorderValues.Single, Size = 4 },
                        new W.BottomBorder           { Val = W.BorderValues.Single, Size = 4 },
                        new W.LeftBorder             { Val = W.BorderValues.Single, Size = 4 },
                        new W.RightBorder            { Val = W.BorderValues.Single, Size = 4 },
                        new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4 },
                        new W.InsideVerticalBorder   { Val = W.BorderValues.Single, Size = 4 })));

                var headerRow = new W.TableRow();
                foreach (var col in new[] { "Flag", "Country", "ISO2", "ISO3" })
                    headerRow.AppendChild(new W.TableCell(
                        new W.Paragraph(new W.Run(new W.RunProperties(new W.Bold()), new W.Text(col)))));
                table.AppendChild(headerRow);

                int rowNum = 0;
                int flagCount = 0;
                uint drawingId = 1;
                foreach (var c in countries)
                {
                    var tableRow = new W.TableRow();
                    tableRow.AppendChild(CreateFlagCell(mainPart, c, ref drawingId, ref flagCount));
                    tableRow.AppendChild(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(c.Country)))));
                    tableRow.AppendChild(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(c.Iso2)))));
                    tableRow.AppendChild(new W.TableCell(new W.Paragraph(new W.Run(new W.Text(c.Iso3)))));
                    table.AppendChild(tableRow);
                    rowNum++;
                    if (rowNum % 50 == 0)
                        log($"  Written {rowNum}/{countries.Count} rows...");
                }

                body.AppendChild(table);
                log($"  Embedded {flagCount} flag images.");
                log("  Saving document...");
                mainPart.Document.Save();
                log($"  Saved: {filePath}");
            });
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
                JsonWriters.MsWordJsonWriter.WriteDocumentFileToJson(sourcePath, outputJsonPath, null, log));
        }

        // ── Flag cell helper ─────────────────────────────────────────────────

        private static W.TableCell CreateFlagCell(
            MainDocumentPart mainPart, MsOfficeCountry_Row country, ref uint drawingId, ref int flagCount)
        {
            if (country.FlagPng is not { Length: > 0 })
                return new W.TableCell(new W.Paragraph());

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(country.FlagPng))
                imagePart.FeedData(stream);

            var widthEmu = PngImage_Helper.GetWidthForHeight(country.FlagPng, FlagHeightEmu);
            var drawing  = OpenXmlFlagImage_Helper.CreateWordInlineDrawing(
                mainPart.GetIdOfPart(imagePart), drawingId++, $"Flag {country.Iso2}", widthEmu, FlagHeightEmu);

            flagCount++;
            return new W.TableCell(new W.Paragraph(new W.Run(drawing)));
        }
    }
}
