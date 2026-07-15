using System;
using System.Collections.Generic;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
using Range = NetOffice.ExcelApi.Range;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entity graph from a running
    // Excel workbook (via the strongly-typed NetOffice object model) and writes it
    // with MsOfficeJsonSerializer. The output matches MSO_MsExcel_JsonWriter.bas.
    // The Excel schema currently has no image/blob properties; the writer options are
    // accepted for signature parity with the other writers.
    public static class MsExcelJsonWriter
    {
        public static void WriteWorkbookToJsonFile(Workbook workbook, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            log?.Invoke($"Writing JSON: {outputFilePath}");

            XL.Workbook entity = BuildWorkbook(workbook, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static XL.Workbook BuildWorkbook(Workbook wb, Action<string>? log)
        {
            var entity = new XL.Workbook
            {
                BuiltInDocumentProperties = Obj<List<XL.BuiltInDocumentProperty>>(() =>
                {
                    var list = new List<XL.BuiltInDocumentProperty>();
                    foreach (Office.DocumentProperty property in (Office.DocumentProperties)wb.BuiltinDocumentProperties)
                    {
                        list.Add(new XL.BuiltInDocumentProperty
                        {
                            Name = Str(() => property.Name),
                            Creator = Long(() => property.Creator),
                            LinkSource = Str(() => property.LinkSource),
                            LinkToContent = Int(() => property.LinkToContent),
                            Type = Int(() => property.Type),
                            Value = Str(() => property.Value),
                        });
                    }
                    return list;
                }) ?? new List<XL.BuiltInDocumentProperty>(),
                CustomDocumentProperties = Obj<List<XL.CustomDocumentProperty>>(() =>
                {
                    var list = new List<XL.CustomDocumentProperty>();
                    foreach (Office.DocumentProperty property in (Office.DocumentProperties)wb.CustomDocumentProperties)
                    {
                        list.Add(new XL.CustomDocumentProperty
                        {
                            Name = Str(() => property.Name),
                            Creator = Long(() => property.Creator),
                            LinkSource = Str(() => property.LinkSource),
                            LinkToContent = Int(() => property.LinkToContent),
                            Type = Int(() => property.Type),
                            Value = Str(() => property.Value),
                        });
                    }
                    return list;
                }) ?? new List<XL.CustomDocumentProperty>(),
                Sheets = Obj<List<XL.Sheet>>(() =>
                {
                    var list = new List<XL.Sheet>();
                    int index = 0;
                    int count = wb.Sheets.Count;
                    foreach (object sheet in wb.Sheets)
                    {
                        index++;
                        log?.Invoke($"  Sheet {index}/{count}...");
                        list.Add(BuildSheet(sheet));
                    }
                    return list;
                }) ?? new List<XL.Sheet>(),
                Name = Str(() => wb.Name),
                FullName = Str(() => wb.FullName),
                Path = Str(() => wb.Path),
                Saved = Bool(() => wb.Saved),
                ReadOnly = Bool(() => wb.ReadOnly),
                HasVBProject = Bool(() => wb.HasVBProject),
                FileFormat = Int(() => wb.FileFormat),
                CodeName = Str(() => wb.CodeName),
            };

            return entity;
        }

        private static XL.Sheet BuildSheet(object sheet)
        {
            // The NetOffice factory materializes each sheet as its concrete wrapper
            // (Worksheet, Chart, DialogSheet, ...), so the wrapper type name matches
            // the VBA writer's TypeName(oSheet).
            return sheet switch
            {
                Worksheet worksheet => BuildWorksheet(worksheet),
                Chart chart => BuildChartSheet(chart),
                _ => BuildOtherSheet(sheet),
            };
        }

        private static XL.Sheet BuildWorksheet(Worksheet worksheet)
        {
            var entity = new XL.Sheet
            {
                Name = Str(() => worksheet.Name),
                Index = Long(() => worksheet.Index),
                CodeName = Str(() => worksheet.CodeName),
                Visible = Int(() => worksheet.Visible),
                StandardWidth = SingleStr(() => worksheet.StandardWidth),
                Type = "Worksheet",
            };

            try
            {
                Range usedRange = worksheet.UsedRange;
                long rowCount = usedRange.Rows.Count;
                long columnCount = usedRange.Columns.Count;

                entity.UsedRange = new XL.UsedRange
                {
                    Address = Str(() => usedRange.Address),
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                };

                entity.Rows = new List<XL.Row>();

                for (long rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    Range row = usedRange.Rows[rowIndex];
                    var rowEntity = new XL.Row { RowIndex = rowIndex, Cells = new List<XL.Cell>() };

                    for (long columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                    {
                        rowEntity.Cells.Add(BuildCell(row.Cells[1, columnIndex], rowIndex, columnIndex));
                    }

                    entity.Rows.Add(rowEntity);
                }
            }
            catch
            {
                entity.UsedRange = null;
                entity.Rows = entity.Rows ?? new List<XL.Row>();
            }

            return entity;
        }

        private static XL.Sheet BuildChartSheet(Chart chart)
        {
            return new XL.Sheet
            {
                Name = Str(() => chart.Name),
                Index = Long(() => chart.Index),
                CodeName = Str(() => chart.CodeName),
                Visible = Int(() => chart.Visible),
                Type = "Chart",
            };
        }

        // Legacy dialog and macro sheets have no dedicated typed path; read the common
        // sheet properties late-bound, exactly like the VBA writer.
        private static XL.Sheet BuildOtherSheet(object sheet)
        {
            dynamic sx = sheet is NetOffice.ICOMObject comObject ? comObject.UnderlyingObject : sheet;

            return new XL.Sheet
            {
                Name = Str(() => sx.Name),
                Index = Long(() => sx.Index),
                CodeName = Str(() => sx.CodeName),
                Visible = Int(() => sx.Visible),
                Type = sheet.GetType().Name,
            };
        }

        private static XL.Cell BuildCell(Range cell, long rowIndex, long columnIndex)
        {
            var entity = new XL.Cell
            {
                Font = Obj(() => new XL.Font
                {
                    Name = Str(() => cell.Font.Name),
                    Size = SingleStr(() => cell.Font.Size),
                    Bold = Bool(() => cell.Font.Bold),
                    Italic = Bool(() => cell.Font.Italic),
                    Underline = Int(() => cell.Font.Underline),
                    Color = Long(() => cell.Font.Color),
                    ColorIndex = Int(() => cell.Font.ColorIndex),
                    StrikeThrough = Bool(() => cell.Font.Strikethrough),
                }),
                Interior = Obj(() => new XL.Interior
                {
                    Color = Long(() => cell.Interior.Color),
                    ColorIndex = Int(() => cell.Interior.ColorIndex),
                    Pattern = Int(() => cell.Interior.Pattern),
                }),
                RowIndex = rowIndex,
                ColumnIndex = columnIndex,
                Address = Str(() => cell.Address),
                Value = Str(() => cell.Value),
                Formula = Str(() => cell.Formula),
                NumberFormat = Str(() => cell.NumberFormat),
                Text = Str(() => cell.Text),
                HorizontalAlignment = Int(() => cell.HorizontalAlignment),
                VerticalAlignment = Int(() => cell.VerticalAlignment),
                WrapText = Bool(() => cell.WrapText),
                MergeCells = Bool(() => cell.MergeCells),
            };

            try
            {
                Comment comment = cell.Comment;
                if (comment != null)
                {
                    entity.HasComment = true;
                    entity.Comment = Str(() => comment.Text());
                }
                else
                {
                    entity.HasComment = false;
                }
            }
            catch
            {
                entity.HasComment = false;
            }

            return entity;
        }
    }
}
