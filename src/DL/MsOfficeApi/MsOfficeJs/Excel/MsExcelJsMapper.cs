using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel
{
    // Maps the Office.js-shaped rows to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel
    // entities and back. VBA properties the Office.js object model cannot provide are set
    // to "**Undefined" (strings), -99 (enums/numerics) or null (booleans/whole objects).
    public static class MsExcelJsMapper
    {
        public static XL.Workbook ToWorkbook(ExcelWorkbookJs_Row source, string? fileName)
        {
            var workbook = new XL.Workbook
            {
                BuiltInDocumentProperties = ToBuiltInDocumentProperties(source.Properties),
                CustomDocumentProperties = new List<XL.CustomDocumentProperty>(),
                Sheets = new List<XL.Sheet>(),
                Name = !string.IsNullOrEmpty(fileName) ? fileName : (source.Name ?? Undefined.String),
                FullName = Undefined.String,
                Path = Undefined.String,
                Saved = null,
                ReadOnly = null,
                HasVBProject = null,
                FileFormat = Undefined.Number,
                CodeName = Undefined.String,
            };

            foreach (var sheet in source.Sheets)
            {
                workbook.Sheets.Add(ToSheet(sheet));
            }

            return workbook;
        }

        private static List<XL.BuiltInDocumentProperty> ToBuiltInDocumentProperties(ExcelDocumentPropertiesJs_Row? properties)
        {
            XL.BuiltInDocumentProperty Property(string name, string? value, int type = 4)
            {
                return new XL.BuiltInDocumentProperty
                {
                    Name = name,
                    Creator = Undefined.NumberLong,
                    LinkSource = null,
                    LinkToContent = 0,
                    Type = type,
                    Value = value,
                };
            }

            return new List<XL.BuiltInDocumentProperty>
            {
                Property("Title", properties?.Title),
                Property("Subject", properties?.Subject),
                Property("Author", properties?.Author),
                Property("Keywords", properties?.Keywords),
                Property("Comments", properties?.Comments),
                Property("Last author", properties?.LastAuthor),
                Property("Revision number", properties?.RevisionNumber),
                Property("Creation date", properties?.CreationDate, 3),
                Property("Last save time", null, 3),
                Property("Category", properties?.Category),
                Property("Content status", null),
            };
        }

        private static XL.Sheet ToSheet(ExcelSheetJs_Row source)
        {
            var sheet = new XL.Sheet
            {
                Name = source.Name ?? string.Empty,
                Index = source.Position + 1,
                CodeName = Undefined.String,
                Visible = source.Visibility?.ToLowerInvariant() switch
                {
                    null or "visible" => -1, // xlSheetVisible
                    "hidden" => 0, // xlSheetHidden
                    "veryhidden" => 2, // xlSheetVeryHidden
                    _ => Undefined.Number,
                },
                StandardWidth = Undefined.String,
                Type = "Worksheet",
            };

            if (source.RowCount <= 0 || source.ColumnCount <= 0)
            {
                sheet.UsedRange = null;
                sheet.Rows = new List<XL.Row>();
                return sheet;
            }

            sheet.UsedRange = new XL.UsedRange
            {
                Address = AbsoluteAddress(source.Address),
                RowCount = source.RowCount,
                ColumnCount = source.ColumnCount,
            };

            sheet.Rows = new List<XL.Row>();

            for (long rowIndex = 1; rowIndex <= source.RowCount; rowIndex++)
            {
                var row = new XL.Row { RowIndex = rowIndex, Cells = new List<XL.Cell>() };

                for (long columnIndex = 1; columnIndex <= source.ColumnCount; columnIndex++)
                {
                    string? value = GetAt(source.Values, rowIndex, columnIndex);
                    string? formula = GetAt(source.Formulas, rowIndex, columnIndex);

                    row.Cells.Add(new XL.Cell
                    {
                        Font = null,
                        Interior = null,
                        RowIndex = rowIndex,
                        ColumnIndex = columnIndex,
                        Address = "$" + ColumnNameFromIndex(columnIndex) + "$" + rowIndex.ToString(CultureInfo.InvariantCulture),
                        Value = value ?? string.Empty,
                        Formula = formula ?? value ?? string.Empty,
                        NumberFormat = Undefined.String,
                        Text = Undefined.String,
                        HorizontalAlignment = Undefined.Number,
                        VerticalAlignment = Undefined.Number,
                        WrapText = null,
                        MergeCells = null,
                        HasComment = null,
                        Comment = null,
                    });
                }

                sheet.Rows.Add(row);
            }

            return sheet;
        }

        private static string? GetAt(List<List<string?>>? values, long rowIndex, long columnIndex)
        {
            if (values == null) return null;
            if (rowIndex - 1 >= values.Count) return null;

            var row = values[(int)(rowIndex - 1)];
            if (row == null || columnIndex - 1 >= row.Count) return null;

            return row[(int)(columnIndex - 1)];
        }

        // "Sheet1!A1:C5" -> "$A$1:$C$5" (matching the VBA UsedRange.Address format).
        private static string AbsoluteAddress(string? address)
        {
            if (string.IsNullOrEmpty(address)) return Undefined.String;

            string range = address!;
            int separator = range.LastIndexOf('!');
            if (separator >= 0) range = range.Substring(separator + 1);

            static string Absolute(string reference)
            {
                string column = new string(reference.TakeWhile(char.IsLetter).ToArray());
                string row = reference.Substring(column.Length);
                return "$" + column + "$" + row;
            }

            return string.Join(":", range.Split(':').Select(Absolute));
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

        public static ExcelWorkbookJs_Row FromWorkbook(XL.Workbook workbook)
        {
            var row = new ExcelWorkbookJs_Row
            {
                Name = workbook.Name,
                Sheets = new List<ExcelSheetJs_Row>(),
            };

            foreach (var sheet in workbook.Sheets ?? new List<XL.Sheet>())
            {
                var sheetRow = new ExcelSheetJs_Row
                {
                    Name = sheet.Name,
                    Position = (int)((sheet.Index ?? 1) - 1),
                    Visibility = sheet.Visible switch
                    {
                        -1 => "Visible",
                        0 => "Hidden",
                        2 => "VeryHidden",
                        _ => null,
                    },
                    Address = sheet.UsedRange?.Address?.Replace("$", string.Empty),
                    RowCount = sheet.UsedRange?.RowCount ?? 0,
                    ColumnCount = sheet.UsedRange?.ColumnCount ?? 0,
                    Values = new List<List<string?>>(),
                    Formulas = new List<List<string?>>(),
                };

                foreach (var dataRow in sheet.Rows ?? new List<XL.Row>())
                {
                    sheetRow.Values.Add((dataRow.Cells ?? new List<XL.Cell>()).Select(c => c.Value).ToList());
                    sheetRow.Formulas.Add((dataRow.Cells ?? new List<XL.Cell>()).Select(c => c.Formula).ToList());
                }

                row.Sheets.Add(sheetRow);
            }

            return row;
        }
    }
}
