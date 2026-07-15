using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters.OpenXmlJsonValues;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entity graph from a .xlsx file
    // and writes it with MsOfficeJsonSerializer. The schema matches
    // MSO_MsExcel_JsonWriter.bas; VBA properties the OpenXML file format cannot provide
    // are written as "**Undefined" (strings), -99 (enums/numerics) or null (booleans).
    public static class MsExcelJsonWriter
    {
        // XlColorIndex / XlPattern / XlUnderlineStyle constants used in the mappings.
        private const int xlColorIndexAutomatic = -4105;
        private const int xlColorIndexNone = -4142;
        private const int xlPatternNone = -4142;
        private const int xlPatternSolid = 1;
        private const int xlUnderlineStyleNone = -4142;

        // The built-in SpreadsheetML number formats (ECMA-376 part 1, §18.8.30).
        private static readonly Dictionary<uint, string> BuiltInNumberFormats = new Dictionary<uint, string>
        {
            [0] = "General", [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00",
            [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
            [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy",
            [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
            [22] = "m/d/yyyy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
        };

        // The legacy indexed color palette (ECMA-376 part 1, §18.8.27); COM ColorIndex n
        // is indexed color n + 7.
        private static readonly string[] IndexedColorPalette =
        {
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
            "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
            "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
            "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
            "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
            "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
        };

        public static void WriteWorkbookFileToJson(string sourcePath, string outputJsonPath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            log?.Invoke($"Writing (OpenXml): {outputJsonPath}");

            using var document = SpreadsheetDocument.Open(sourcePath, false);

            XL.Workbook entity = BuildWorkbook(document, sourcePath, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputJsonPath);

            log?.Invoke($"Done: {outputJsonPath}");
        }

        private static XL.Workbook BuildWorkbook(SpreadsheetDocument document, string sourcePath, Action<string>? log)
        {
            var workbookPart = document.WorkbookPart
                ?? throw new InvalidOperationException("Cannot read workbook part.");

            var entity = new XL.Workbook
            {
                BuiltInDocumentProperties = OpenXmlDocumentProperties.BuildBuiltIn(document.PackageProperties,
                    document.ExtendedFilePropertiesPart,
                    (name, value, type) => new XL.BuiltInDocumentProperty
                    {
                        Name = name,
                        Creator = Undefined.NumberLong,
                        LinkSource = null,
                        LinkToContent = 0,
                        Type = type,
                        Value = value,
                    }),
                CustomDocumentProperties = OpenXmlDocumentProperties.BuildCustom(document.CustomFilePropertiesPart,
                    (name, value, type) => new XL.CustomDocumentProperty
                    {
                        Name = name,
                        Creator = Undefined.NumberLong,
                        LinkSource = null,
                        LinkToContent = 0,
                        Type = type,
                        Value = value,
                    }),
                Sheets = BuildSheets(workbookPart, log),
                Name = Path.GetFileName(sourcePath),
                FullName = sourcePath,
                Path = Path.GetDirectoryName(sourcePath) ?? string.Empty,
                Saved = null,
                ReadOnly = null,
                HasVBProject = workbookPart.VbaProjectPart != null,
                FileFormat = Path.GetExtension(sourcePath).ToLowerInvariant() switch
                {
                    ".xlsx" => 51, // xlOpenXMLWorkbook
                    ".xlsm" => 52, // xlOpenXMLWorkbookMacroEnabled
                    _ => Undefined.Number,
                },
                CodeName = workbookPart.Workbook?.WorkbookProperties?.CodeName?.Value ?? Undefined.String,
            };

            return entity;
        }

        private static List<XL.Sheet> BuildSheets(WorkbookPart workbookPart, Action<string>? log)
        {
            var list = new List<XL.Sheet>();

            var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

            long index = 0;
            foreach (var sheet in sheets)
            {
                index++;
                log?.Invoke($"  Sheet {index}/{sheets.Count}...");

                string? relationshipId = sheet.Id?.Value;
                var part = relationshipId == null ? null : workbookPart.GetPartById(relationshipId);

                int visible = sheet.State?.Value switch
                {
                    null => -1, // xlSheetVisible
                    SheetStateValues state when state == SheetStateValues.Visible => -1,
                    SheetStateValues state when state == SheetStateValues.Hidden => 0,
                    SheetStateValues state when state == SheetStateValues.VeryHidden => 2,
                    _ => Undefined.Number,
                };

                if (part is WorksheetPart worksheetPart)
                {
                    var entity = new XL.Sheet
                    {
                        Name = sheet.Name?.Value ?? string.Empty,
                        Index = index,
                        CodeName = worksheetPart.Worksheet?.SheetProperties?.CodeName?.Value ?? Undefined.String,
                        Visible = visible,
                        StandardWidth = worksheetPart.Worksheet?.SheetFormatProperties?.DefaultColumnWidth?.Value is double width
                            ? ((float)width).ToString(CultureInfo.InvariantCulture)
                            : Undefined.String,
                        Type = "Worksheet",
                    };

                    BuildUsedRangeAndRows(entity, worksheetPart, workbookPart, sharedStrings);

                    list.Add(entity);
                }
                else
                {
                    list.Add(new XL.Sheet
                    {
                        UsedRange = null,
                        Rows = null,
                        Name = sheet.Name?.Value ?? string.Empty,
                        Index = index,
                        CodeName = Undefined.String,
                        Visible = visible,
                        StandardWidth = null,
                        Type = part is ChartsheetPart ? "Chart" : (part?.GetType().Name ?? "Unknown"),
                    });
                }
            }

            return list;
        }

        private static void BuildUsedRangeAndRows(XL.Sheet entity, WorksheetPart worksheetPart, WorkbookPart workbookPart,
            SharedStringTable? sharedStrings)
        {
            var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();

            var rows = sheetData?.Elements<Row>().ToList() ?? new List<Row>();

            long rowCount = 0;
            long columnCount = 0;

            var cellsByRow = new Dictionary<long, Dictionary<long, Cell>>();

            foreach (var row in rows)
            {
                long rowIndex = (long?)row.RowIndex?.Value ?? 0;
                if (rowIndex == 0) continue;

                var rowCells = new Dictionary<long, Cell>();

                foreach (var cell in row.Elements<Cell>())
                {
                    long columnIndex = ColumnIndexFromReference(cell.CellReference?.Value);
                    if (columnIndex == 0) continue;

                    rowCells[columnIndex] = cell;
                    if (columnIndex > columnCount) columnCount = columnIndex;
                }

                cellsByRow[rowIndex] = rowCells;
                if (rowIndex > rowCount) rowCount = rowIndex;
            }

            // Match the VBA writer address format ($A$1:$C$5).
            string? dimension = worksheetPart.Worksheet?.SheetDimension?.Reference?.Value;

            entity.UsedRange = new XL.UsedRange
            {
                Address = dimension == null
                    ? Undefined.String
                    : string.Join(":", dimension.Split(':').Select(AbsoluteReference)),
                RowCount = rowCount,
                ColumnCount = columnCount,
            };

            var mergedRanges = worksheetPart.Worksheet?.GetFirstChild<MergeCells>()?.Elements<MergeCell>()
                .Select(m => m.Reference?.Value)
                .Where(r => r != null)
                .Cast<string>()
                .ToList() ?? new List<string>();

            var comments = worksheetPart.WorksheetCommentsPart?.Comments;

            entity.Rows = new List<XL.Row>();

            for (long rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                var rowEntity = new XL.Row { RowIndex = rowIndex, Cells = new List<XL.Cell>() };

                cellsByRow.TryGetValue(rowIndex, out var rowCells);

                for (long columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    Cell? cell = null;
                    rowCells?.TryGetValue(columnIndex, out cell);

                    rowEntity.Cells.Add(BuildCell(cell, rowIndex, columnIndex, workbookPart, sharedStrings, mergedRanges, comments));
                }

                entity.Rows.Add(rowEntity);
            }
        }

        private static XL.Cell BuildCell(Cell? cell, long rowIndex, long columnIndex, WorkbookPart workbookPart,
            SharedStringTable? sharedStrings, List<string> mergedRanges, Comments? comments)
        {
            string reference = ColumnNameFromIndex(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);
            string address = "$" + ColumnNameFromIndex(columnIndex) + "$" + rowIndex.ToString(CultureInfo.InvariantCulture);

            string? value = cell == null ? string.Empty : GetCellValue(cell, sharedStrings);
            string? formula = cell?.CellFormula?.Text ?? string.Empty;
            if (formula.Length > 0) formula = "=" + formula;
            else formula = value ?? string.Empty;

            string? commentText = comments?.CommentList?.Elements<Comment>()
                .FirstOrDefault(c => c.Reference?.Value == reference)?
                .CommentText?.InnerText;

            var (font, interior, numberFormat, wrapText, horizontalAlignment, verticalAlignment) =
                GetCellFormat(cell, workbookPart);

            return new XL.Cell
            {
                Font = font,
                Interior = interior,
                RowIndex = rowIndex,
                ColumnIndex = columnIndex,
                Address = address,
                Value = value,
                Formula = formula,
                NumberFormat = numberFormat,
                Text = Undefined.String,
                HorizontalAlignment = horizontalAlignment,
                VerticalAlignment = verticalAlignment,
                WrapText = wrapText,
                MergeCells = mergedRanges.Any(range => ReferenceIsInRange(reference, range)),
                HasComment = commentText != null,
                Comment = commentText,
            };
        }

        private static string? GetCellValue(Cell cell, SharedStringTable? sharedStrings)
        {
            string? raw = cell.CellValue?.Text;

            if (cell.DataType?.Value == CellValues.SharedString && raw != null && sharedStrings != null)
            {
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                    && index >= 0 && index < sharedStrings.ChildElements.Count)
                {
                    return sharedStrings.ChildElements[index].InnerText;
                }
            }

            if (cell.DataType?.Value == CellValues.InlineString)
            {
                return cell.InlineString?.InnerText ?? raw;
            }

            if (cell.DataType?.Value == CellValues.Boolean && raw != null)
            {
                return raw == "1" ? "True" : "False";
            }

            return raw ?? string.Empty;
        }

        private static (XL.Font?, XL.Interior?, string?, bool?, int, int) GetCellFormat(Cell? cell, WorkbookPart workbookPart)
        {
            const int xlHAlignGeneral = 1;
            const int xlVAlignBottom = -4107;

            try
            {
                var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
                if (stylesheet == null)
                {
                    return (null, null, "General", null, xlHAlignGeneral, xlVAlignBottom);
                }

                // A cell without an explicit style index uses the default cell format (xf 0).
                int styleIndex = (int)(cell?.StyleIndex?.Value ?? 0);
                var cellFormat = (CellFormat?)stylesheet.CellFormats?.ChildElements[styleIndex];
                if (cellFormat == null)
                {
                    return (null, null, "General", null, xlHAlignGeneral, xlVAlignBottom);
                }

                XL.Font? fontEntity = null;
                if (cellFormat.FontId?.Value != null)
                {
                    var font = (Font?)stylesheet.Fonts?.ChildElements[(int)cellFormat.FontId.Value];
                    if (font != null)
                    {
                        var (color, colorIndex) = OleColorAndIndexFrom(font.Color);

                        fontEntity = new XL.Font
                        {
                            Name = font.FontName?.Val?.Value ?? Undefined.String,
                            Size = font.FontSize?.Val?.Value is double size
                                ? ((float)size).ToString(CultureInfo.InvariantCulture)
                                : Undefined.String,
                            Bold = font.Bold != null && (font.Bold.Val?.Value ?? true),
                            Italic = font.Italic != null && (font.Italic.Val?.Value ?? true),
                            Underline = XlUnderlineStyleFrom(font.Underline),
                            Color = color,
                            ColorIndex = colorIndex,
                            StrikeThrough = font.Strike != null && (font.Strike.Val?.Value ?? true),
                        };
                    }
                }

                XL.Interior? interiorEntity = null;
                if (cellFormat.FillId?.Value != null)
                {
                    var fill = (Fill?)stylesheet.Fills?.ChildElements[(int)cellFormat.FillId.Value];
                    interiorEntity = BuildInterior(fill);
                }

                string numberFormat = "General";
                if (cellFormat.NumberFormatId?.Value is uint numberFormatId && numberFormatId > 0)
                {
                    var custom = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
                        .FirstOrDefault(f => f.NumberFormatId?.Value == numberFormatId);
                    numberFormat = custom?.FormatCode?.Value
                        ?? (BuiltInNumberFormats.TryGetValue(numberFormatId, out string? builtIn)
                            ? builtIn
                            : numberFormatId.ToString(CultureInfo.InvariantCulture));
                }

                var alignment = cellFormat.Alignment;
                bool wrapText = alignment?.WrapText?.Value ?? false;

                int horizontalAlignment = alignment?.Horizontal?.Value switch
                {
                    null => xlHAlignGeneral,
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.General => 1,
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Left => -4131, // xlHAlignLeft
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Center => -4108, // xlHAlignCenter
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.CenterContinuous => 7, // xlHAlignCenterAcrossSelection
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Right => -4152, // xlHAlignRight
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Fill => 5, // xlHAlignFill
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Justify => -4130, // xlHAlignJustify
                    HorizontalAlignmentValues align when align == HorizontalAlignmentValues.Distributed => -4117, // xlHAlignDistributed
                    _ => Undefined.Number,
                };

                int verticalAlignment = alignment?.Vertical?.Value switch
                {
                    null => xlVAlignBottom,
                    VerticalAlignmentValues align when align == VerticalAlignmentValues.Top => -4160, // xlVAlignTop
                    VerticalAlignmentValues align when align == VerticalAlignmentValues.Center => -4108, // xlVAlignCenter
                    VerticalAlignmentValues align when align == VerticalAlignmentValues.Bottom => -4107, // xlVAlignBottom
                    VerticalAlignmentValues align when align == VerticalAlignmentValues.Justify => -4130, // xlVAlignJustify
                    VerticalAlignmentValues align when align == VerticalAlignmentValues.Distributed => -4117, // xlVAlignDistributed
                    _ => Undefined.Number,
                };

                return (fontEntity, interiorEntity, numberFormat, wrapText, horizontalAlignment, verticalAlignment);
            }
            catch
            {
                return (null, null, "General", null, xlHAlignGeneral, xlVAlignBottom);
            }
        }

        private static XL.Interior BuildInterior(Fill? fill)
        {
            var patternFill = fill?.PatternFill;
            var patternType = patternFill?.PatternType?.Value;

            if (patternFill == null || patternType == null || patternType == PatternValues.None)
            {
                // COM reports unfilled cells as white with no color index.
                return new XL.Interior
                {
                    Color = 16777215,
                    ColorIndex = xlColorIndexNone,
                    Pattern = xlPatternNone,
                };
            }

            var (color, colorIndex) = OleColorAndIndexFrom(patternFill.ForegroundColor);

            return new XL.Interior
            {
                Color = color,
                ColorIndex = colorIndex,
                Pattern = XlPatternFrom(patternType.Value),
            };
        }

        private static int XlPatternFrom(PatternValues patternType)
        {
            // XlPattern values.
            if (patternType == PatternValues.Solid) return xlPatternSolid;
            if (patternType == PatternValues.MediumGray) return -4125; // xlPatternGray50
            if (patternType == PatternValues.DarkGray) return -4126; // xlPatternGray75
            if (patternType == PatternValues.LightGray) return -4124; // xlPatternGray25
            if (patternType == PatternValues.Gray125) return 17; // xlPatternGray16
            if (patternType == PatternValues.Gray0625) return 18; // xlPatternGray8
            if (patternType == PatternValues.DarkHorizontal) return -4128; // xlPatternHorizontal
            if (patternType == PatternValues.DarkVertical) return -4166; // xlPatternVertical
            if (patternType == PatternValues.DarkDown) return -4121; // xlPatternDown
            if (patternType == PatternValues.DarkUp) return -4162; // xlPatternUp
            if (patternType == PatternValues.DarkGrid) return 9; // xlPatternChecker
            if (patternType == PatternValues.DarkTrellis) return 10; // xlPatternSemiGray75
            if (patternType == PatternValues.LightHorizontal) return 11; // xlPatternLightHorizontal
            if (patternType == PatternValues.LightVertical) return 12; // xlPatternLightVertical
            if (patternType == PatternValues.LightDown) return 13; // xlPatternLightDown
            if (patternType == PatternValues.LightUp) return 14; // xlPatternLightUp
            if (patternType == PatternValues.LightGrid) return 15; // xlPatternGrid
            if (patternType == PatternValues.LightTrellis) return 16; // xlPatternCrissCross

            return Undefined.Number;
        }

        private static int XlUnderlineStyleFrom(Underline? underline)
        {
            if (underline == null) return xlUnderlineStyleNone;

            return underline.Val?.Value switch
            {
                null => 2, // xlUnderlineStyleSingle (the element default)
                UnderlineValues value when value == UnderlineValues.None => xlUnderlineStyleNone,
                UnderlineValues value when value == UnderlineValues.Single => 2,
                UnderlineValues value when value == UnderlineValues.Double => -4119, // xlUnderlineStyleDouble
                UnderlineValues value when value == UnderlineValues.SingleAccounting => 4,
                UnderlineValues value when value == UnderlineValues.DoubleAccounting => 5,
                _ => Undefined.Number,
            };
        }

        // Maps a styles.xml color to the COM (Color, ColorIndex) pair. Indexed colors are
        // resolved through the legacy palette; theme colors are not resolved (the tint
        // math would only approximate what Excel reports).
        private static (long?, int) OleColorAndIndexFrom(ColorType? color)
        {
            if (color == null) return (null, xlColorIndexAutomatic);

            if (color.Rgb?.Value is string argb)
            {
                return (OleColorFromHex(TrimArgb(argb)), Undefined.Number);
            }

            if (color.Indexed?.Value is uint indexed)
            {
                string? hex = indexed < IndexedColorPalette.Length ? IndexedColorPalette[indexed] : null;
                int colorIndex = indexed >= 8 && indexed <= 63 ? (int)indexed - 7 : xlColorIndexAutomatic;

                return (OleColorFromHex(hex), colorIndex);
            }

            if (color.Auto?.Value == true) return (null, xlColorIndexAutomatic);

            return (null, Undefined.Number);
        }

        private static string? TrimArgb(string? argb)
        {
            if (string.IsNullOrEmpty(argb)) return null;

            return argb!.Length == 8 ? argb.Substring(2) : argb;
        }

        private static long ColumnIndexFromReference(string? reference)
        {
            if (string.IsNullOrEmpty(reference)) return 0;

            long index = 0;
            foreach (char c in reference!)
            {
                if (!char.IsLetter(c)) break;
                index = index * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
            }

            return index;
        }

        private static string ColumnNameFromIndex(long index)
        {
            string name = string.Empty;

            while (index > 0)
            {
                long remainder = (index - 1) % 26;
                name = (char)('A' + remainder) + name;
                index = (index - 1) / 26;
            }

            return name;
        }

        private static string AbsoluteReference(string reference)
        {
            string column = new string(reference.TakeWhile(char.IsLetter).ToArray());
            string row = reference.Substring(column.Length);

            return "$" + column + "$" + row;
        }

        private static bool ReferenceIsInRange(string reference, string range)
        {
            var parts = range.Split(':');
            if (parts.Length != 2) return false;

            long column = ColumnIndexFromReference(reference);
            long row = RowFromReference(reference);

            long fromColumn = ColumnIndexFromReference(parts[0]);
            long fromRow = RowFromReference(parts[0]);
            long toColumn = ColumnIndexFromReference(parts[1]);
            long toRow = RowFromReference(parts[1]);

            return column >= fromColumn && column <= toColumn && row >= fromRow && row <= toRow;
        }

        private static long RowFromReference(string reference)
        {
            string digits = new string(reference.SkipWhile(char.IsLetter).ToArray());

            return long.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out long row) ? row : 0;
        }
    }
}
