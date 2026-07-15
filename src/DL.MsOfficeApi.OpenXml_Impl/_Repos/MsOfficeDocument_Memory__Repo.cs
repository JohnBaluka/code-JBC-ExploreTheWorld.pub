using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    public class MsOfficeDocument_Memory__Repo : MsOfficeDocument_Memory__Repo__Interface
    {
        private const long FlagHeightEmu = 200000;
        private const double ExcelFlagRowHeightPoints = 20;

        // ── Word ────────────────────────────────────────────────────────────────────

        public byte[] BuildWordBytes(IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            log($"Building Word document ({countries.Count} countries)...");
            var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new W.Document();
                var body = main.Document.AppendChild(new W.Body());

                body.AppendChild(new W.Paragraph(
                    new W.Run(new W.RunProperties(new W.Bold(), new W.FontSize { Val = "32" }),
                              new W.Text("Countries Now — countriesnow.space"))));

                body.AppendChild(new W.Paragraph(
                    new W.Run(new W.Text($"Total countries: {countries.Count}"))));

                var table = body.AppendChild(new W.Table(new W.TableProperties(
                    new W.TableBorders(
                        new W.TopBorder              { Val = W.BorderValues.Single, Size = 4 },
                        new W.BottomBorder           { Val = W.BorderValues.Single, Size = 4 },
                        new W.LeftBorder             { Val = W.BorderValues.Single, Size = 4 },
                        new W.RightBorder            { Val = W.BorderValues.Single, Size = 4 },
                        new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4 },
                        new W.InsideVerticalBorder   { Val = W.BorderValues.Single, Size = 4 }))));

                table.AppendChild(MakeWordRow(new[] { "Flag", "Country", "ISO2", "ISO3" }, isHeader: true));

                uint drawingId = 1;
                foreach (var c in countries)
                {
                    var row = MakeWordRow(new[] { c.Country, c.Iso2, c.Iso3 });
                    row.InsertAt(MakeWordFlagCell(main, c, ref drawingId), 0);
                    table.AppendChild(row);
                }

                main.Document.Save();
            }
            log($"  Word: {countries.Count} rows written");
            return ms.ToArray();
        }

        private static W.TableRow MakeWordRow(string[] cells, bool isHeader = false)
        {
            var row = new W.TableRow();
            foreach (var text in cells)
            {
                var runProps = isHeader ? new W.RunProperties(new W.Bold()) : new W.RunProperties();
                row.AppendChild(new W.TableCell(
                    new W.TableCellProperties(new W.TableCellWidth { Type = W.TableWidthUnitValues.Auto }),
                    new W.Paragraph(new W.Run(runProps, new W.Text(text)))));
            }
            return row;
        }

        private static W.TableCell MakeWordFlagCell(
            MainDocumentPart mainPart, MsOfficeCountry_Row country, ref uint drawingId)
        {
            var cellProps = new W.TableCellProperties(new W.TableCellWidth { Type = W.TableWidthUnitValues.Auto });
            if (country.FlagPng is not { Length: > 0 })
                return new W.TableCell(cellProps, new W.Paragraph());

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(country.FlagPng))
                imagePart.FeedData(stream);

            var widthEmu = PngImage_Helper.GetWidthForHeight(country.FlagPng, FlagHeightEmu);
            var drawing  = OpenXmlFlagImage_Helper.CreateWordInlineDrawing(
                mainPart.GetIdOfPart(imagePart), drawingId++, $"Flag {country.Iso2}", widthEmu, FlagHeightEmu);

            return new W.TableCell(cellProps, new W.Paragraph(new W.Run(drawing)));
        }

        // ── Excel ───────────────────────────────────────────────────────────────────

        public byte[] BuildExcelBytes(IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            log($"Building Excel workbook ({countries.Count} countries)...");
            var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var wbPart    = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                var wsPart    = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();
                var columns   = new S.Columns(
                    new S.Column { Min = 1, Max = 1, Width = 7, CustomWidth = true });
                wsPart.Worksheet = new S.Worksheet(columns, sheetData);

                wbPart.Workbook.AppendChild(new S.Sheets()).AppendChild(new S.Sheet
                {
                    Id      = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1,
                    Name    = "Countries"
                });

                sheetData.AppendChild(MakeExcelRow(new[] { "Flag", "Country", "ISO2", "ISO3" }));

                var worksheetDrawing = new Xdr.WorksheetDrawing();
                DrawingsPart? drawingsPart = null;
                uint drawingId = 1;
                int rowIndex = 1; // 0-based anchor row; row 0 is the header
                foreach (var c in countries)
                {
                    var row = MakeExcelRow(new[] { string.Empty, c.Country, c.Iso2, c.Iso3 });
                    if (c.FlagPng is { Length: > 0 })
                    {
                        row.Height       = ExcelFlagRowHeightPoints;
                        row.CustomHeight = true;

                        drawingsPart ??= wsPart.AddNewPart<DrawingsPart>();
                        var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
                        using (var stream = new MemoryStream(c.FlagPng))
                            imagePart.FeedData(stream);

                        var widthEmu = PngImage_Helper.GetWidthForHeight(c.FlagPng, FlagHeightEmu);
                        worksheetDrawing.AppendChild(OpenXmlFlagImage_Helper.CreateSpreadsheetAnchor(
                            drawingsPart.GetIdOfPart(imagePart), drawingId++, $"Flag {c.Iso2}",
                            columnIndex: 0, rowIndex: rowIndex, widthEmu, FlagHeightEmu));
                    }
                    sheetData.AppendChild(row);
                    rowIndex++;
                }

                if (drawingsPart != null)
                {
                    drawingsPart.WorksheetDrawing = worksheetDrawing;
                    wsPart.Worksheet.AppendChild(new S.Drawing { Id = wsPart.GetIdOfPart(drawingsPart) });
                }

                wbPart.Workbook.Save();
            }
            log($"  Excel: {countries.Count} rows written");
            return ms.ToArray();
        }

        private static S.Row MakeExcelRow(string[] values)
        {
            var row = new S.Row();
            foreach (var v in values)
                row.AppendChild(new S.Cell
                {
                    DataType     = S.CellValues.InlineString,
                    InlineString = new S.InlineString(new S.Text(v))
                });
            return row;
        }

        // ── PowerPoint ──────────────────────────────────────────────────────────────

        public byte[] BuildPowerPointBytes(IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            log($"Building PowerPoint presentation ({countries.Count} countries)...");
            var ms = new MemoryStream();
            int slideCount;
            using (var prs = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                var prsPart     = prs.AddPresentationPart();
                var presentation = prsPart.Presentation = new P.Presentation
                {
                    SlideSize = new P.SlideSize { Cx = 9144000, Cy = 5143500 },
                    NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9144000 }
                };

                var smPart     = prsPart.AddNewPart<SlideMasterPart>();
                var layoutPart = smPart.AddNewPart<SlideLayoutPart>();
                layoutPart.SlideLayout = MakeSlideLayout();
                smPart.SlideMaster     = MakeSlideMaster(smPart.GetIdOfPart(layoutPart));

                // PowerPoint reports the file as corrupted unless the layout references its
                // master back and the master has a theme part (also referenced by the
                // presentation part), like every PowerPoint-saved package.
                layoutPart.AddPart(smPart);
                var themePart = smPart.AddNewPart<ThemePart>();
                themePart.Theme = OpenXmlMinimalTheme.Create();
                themePart.Theme.Save();
                prsPart.AddPart(themePart);

                presentation.SlideMasterIdList = new P.SlideMasterIdList(
                    new P.SlideMasterId { Id = 2147483648U, RelationshipId = prsPart.GetIdOfPart(smPart) });

                presentation.SlideIdList = new P.SlideIdList();

                uint sid = 256U;
                AppendSlide(prsPart, layoutPart, ref sid,
                    "Countries Now API",
                    $"countriesnow.space — {countries.Count} countries");

                const int pageSize = 20;
                const int maxPages = 10;
                for (int page = 0; page < maxPages && page * pageSize < countries.Count; page++)
                {
                    AppendCountrySlide(prsPart, layoutPart, ref sid,
                        $"Countries — page {page + 1}",
                        countries.Skip(page * pageSize).Take(pageSize).ToList());
                }

                presentation.Save();
                slideCount = ((P.SlideIdList)presentation.SlideIdList).Count();
            }
            log($"  PowerPoint: {slideCount} slides ({countries.Count} countries)");
            return ms.ToArray();
        }

        private static P.SlideLayout MakeSlideLayout() =>
            new(new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(
                            new A.TransformGroup(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 0L, Cy = 0L },
                                new A.ChildOffset { X = 0L, Y = 0L },
                                new A.ChildExtents { Cx = 0L, Cy = 0L })))),
                new P.ColorMapOverride(new A.MasterColorMapping()));

        private static P.SlideMaster MakeSlideMaster(string layoutRelId) =>
            new(new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(
                            new A.TransformGroup(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 0L, Cy = 0L },
                                new A.ChildOffset { X = 0L, Y = 0L },
                                new A.ChildExtents { Cx = 0L, Cy = 0L })))),
                new P.ColorMap
                {
                    Background1       = A.ColorSchemeIndexValues.Light1,
                    Text1             = A.ColorSchemeIndexValues.Dark1,
                    Background2       = A.ColorSchemeIndexValues.Light2,
                    Text2             = A.ColorSchemeIndexValues.Dark2,
                    Accent1           = A.ColorSchemeIndexValues.Accent1,
                    Accent2           = A.ColorSchemeIndexValues.Accent2,
                    Accent3           = A.ColorSchemeIndexValues.Accent3,
                    Accent4           = A.ColorSchemeIndexValues.Accent4,
                    Accent5           = A.ColorSchemeIndexValues.Accent5,
                    Accent6           = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink         = A.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new P.SlideLayoutIdList(
                    // ST_SlideLayoutId requires ids >= 0x80000000.
                    new P.SlideLayoutId { Id = 2147483649U, RelationshipId = layoutRelId }));

        private static SlidePart AppendSlidePart(
            PresentationPart prsPart, SlideLayoutPart layoutPart, ref uint slideId, P.ShapeTree shapeTree)
        {
            var slidePart = prsPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(new P.CommonSlideData(shapeTree));

            ((P.SlideIdList)prsPart.Presentation.SlideIdList!).AppendChild(
                new P.SlideId { Id = slideId++, RelationshipId = prsPart.GetIdOfPart(slidePart) });
            return slidePart;
        }

        private static void AppendSlide(
            PresentationPart prsPart, SlideLayoutPart layoutPart,
            ref uint slideId, string title, string body)
        {
            AppendSlidePart(prsPart, layoutPart, ref slideId, new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                MakeTextShape(2U, "Title", title,
                    x: 457200L,  y: 274638L,  cx: 8229600L, cy: 1143000L,
                    fontSize: 2800, bold: true),
                MakeTextShape(3U, "Body", body,
                    x: 457200L,  y: 1600200L, cx: 8229600L, cy: 3657600L,
                    fontSize: 1400, bold: false)));
        }

        // One slide with a flag picture + text line per country row (20 rows per slide).
        private static void AppendCountrySlide(
            PresentationPart prsPart, SlideLayoutPart layoutPart,
            ref uint slideId, string title, IList<MsOfficeCountry_Row> rows)
        {
            const long contentTopEmu = 1600200;
            const long rowHeightEmu  = 182880;
            const long flagLeftEmu   = 457200;
            const long flagHeightEmu = 150000;
            const long textLeftEmu   = 1066800;
            const long textWidthEmu  = 7620000;

            var slidePart = AppendSlidePart(prsPart, layoutPart, ref slideId, new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                MakeTextShape(2U, "Title", title,
                    x: 457200L, y: 274638L, cx: 8229600L, cy: 1143000L,
                    fontSize: 2800, bold: true)));

            var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;
            uint shapeId = 3;
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                long rowTop = contentTopEmu + r * rowHeightEmu;

                if (row.FlagPng is { Length: > 0 })
                {
                    var imagePart = slidePart.AddImagePart(ImagePartType.Png);
                    using (var stream = new MemoryStream(row.FlagPng))
                        imagePart.FeedData(stream);

                    var widthEmu = PngImage_Helper.GetWidthForHeight(row.FlagPng, flagHeightEmu);
                    tree.AppendChild(OpenXmlFlagImage_Helper.CreatePresentationPicture(
                        slidePart.GetIdOfPart(imagePart), shapeId++, $"Flag {row.Iso2}",
                        flagLeftEmu, rowTop + (rowHeightEmu - flagHeightEmu) / 2, widthEmu, flagHeightEmu));
                }

                tree.AppendChild(MakeTextShape(shapeId++, $"Row {r + 1}",
                    $"{row.Country,-35}  {row.Iso2,-4}  {row.Iso3}",
                    x: textLeftEmu, y: rowTop, cx: textWidthEmu, cy: rowHeightEmu,
                    fontSize: 1000, bold: false));
            }
        }

        private static P.Shape MakeTextShape(
            uint id, string name, string text,
            long x, long y, long cx, long cy, int fontSize, bool bold) =>
            new(new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset  { X = x,  Y = y  },
                        new A.Extents { Cx = cx, Cy = cy })),
                new P.TextBody(
                    new A.BodyProperties { Wrap = A.TextWrappingValues.Square },
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US", FontSize = fontSize, Bold = bold },
                            new A.Text(text)))));
    }
}
