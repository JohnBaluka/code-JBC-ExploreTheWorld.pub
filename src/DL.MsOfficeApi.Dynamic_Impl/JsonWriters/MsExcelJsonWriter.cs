using System;
using System.Collections;
using System.Collections.Generic;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel entity graph from a running
    // Excel workbook COM object (late-bound or Interop PIA) and writes it with
    // MsOfficeJsonSerializer. The output matches MSO_MsExcel_JsonWriter.bas.
    // The Excel schema currently has no image/blob properties; the writer options are
    // accepted for signature parity with the other writers.
    public static class MsExcelJsonWriter
    {
        // Accepts any Excel Workbook COM object (late-bound or Interop PIA) —
        // used by the Dynamic repos and the Save-As-JSON helpers.
        public static void WriteWorkbookComToJsonFile(object workbookComObject, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            log?.Invoke($"Writing JSON: {outputFilePath}");

            dynamic com = workbookComObject;

            XL.Workbook entity = BuildWorkbook(com, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static XL.Workbook BuildWorkbook(dynamic wb, Action<string>? log)
        {
            var entity = new XL.Workbook
            {
                BuiltInDocumentProperties = Obj<List<XL.BuiltInDocumentProperty>>(() =>
                {
                    var list = new List<XL.BuiltInDocumentProperty>();
                    foreach (dynamic property in (IEnumerable)wb.BuiltInDocumentProperties)
                    {
                        dynamic px = property;
                        list.Add(new XL.BuiltInDocumentProperty
                        {
                            Name = Str(() => px.Name),
                            Creator = Long(() => px.Creator),
                            LinkSource = Str(() => px.LinkSource),
                            LinkToContent = Int(() => px.LinkToContent),
                            Type = Int(() => px.Type),
                            Value = Str(() => px.Value),
                        });
                    }
                    return list;
                }) ?? new List<XL.BuiltInDocumentProperty>(),
                CustomDocumentProperties = Obj<List<XL.CustomDocumentProperty>>(() =>
                {
                    var list = new List<XL.CustomDocumentProperty>();
                    foreach (dynamic property in (IEnumerable)wb.CustomDocumentProperties)
                    {
                        dynamic px = property;
                        list.Add(new XL.CustomDocumentProperty
                        {
                            Name = Str(() => px.Name),
                            Creator = Long(() => px.Creator),
                            LinkSource = Str(() => px.LinkSource),
                            LinkToContent = Int(() => px.LinkToContent),
                            Type = Int(() => px.Type),
                            Value = Str(() => px.Value),
                        });
                    }
                    return list;
                }) ?? new List<XL.CustomDocumentProperty>(),
                Sheets = Obj<List<XL.Sheet>>(() =>
                {
                    var list = new List<XL.Sheet>();
                    int index = 0;
                    int count = Convert.ToInt32(wb.Sheets.Count);
                    foreach (dynamic sheet in wb.Sheets)
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

        private static XL.Sheet BuildSheet(dynamic sheet)
        {
            string typeName = GetSheetTypeName(sheet);
            bool isWorksheet = typeName == "Worksheet";

            var entity = new XL.Sheet
            {
                Name = Str(() => sheet.Name),
                Index = Long(() => sheet.Index),
                CodeName = Str(() => sheet.CodeName),
                Visible = Int(() => sheet.Visible),
                StandardWidth = isWorksheet ? SingleStr(() => sheet.StandardWidth) : null,
                Type = typeName,
            };

            if (!isWorksheet)
            {
                return entity;
            }

            try
            {
                dynamic usedRange = sheet.UsedRange;
                long rowCount = Convert.ToInt64(usedRange.Rows.Count);
                long columnCount = Convert.ToInt64(usedRange.Columns.Count);

                entity.UsedRange = new XL.UsedRange
                {
                    Address = Str(() => usedRange.Address),
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                };

                entity.Rows = new List<XL.Row>();

                for (long rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    dynamic row = usedRange.Rows[rowIndex];
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

        private static string GetSheetTypeName(dynamic sheet)
        {
            // Worksheets expose UsedRange; chart sheets and macro sheets do not.
            try
            {
                dynamic _ = sheet.UsedRange;
                return "Worksheet";
            }
            catch
            {
                return "Chart";
            }
        }

        private static XL.Cell BuildCell(dynamic cell, long rowIndex, long columnIndex)
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
                dynamic comment = cell.Comment;
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
