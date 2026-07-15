using System;
using System.Collections.Generic;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Range = Microsoft.Office.Interop.Word.Range;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entity graph from a running
    // Word document (via the strongly-typed Microsoft.Office.Interop object model) and writes it
    // with MsOfficeJsonSerializer. The output matches MSO_MsWord_JsonWriter.bas.
    // Like the VBA writer, the InlineShape Image property is always null on this path
    // (the OpenXML writer provides the image bytes); the writer options are accepted
    // for signature parity.
    public static class MsWordJsonWriter
    {
        private const int MaxRunsPerParagraph = 500;

        public static void WriteDocumentToJsonFile(Document document, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            log?.Invoke($"Writing JSON: {outputFilePath}");

            WD.Document entity = BuildDocument(document, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static WD.Document BuildDocument(Document doc, Action<string>? log)
        {
            var entity = new WD.Document
            {
                BuiltInDocumentProperties = Obj<List<WD.BuiltInDocumentProperty>>(() =>
                {
                    var list = new List<WD.BuiltInDocumentProperty>();
                    foreach (Office.DocumentProperty property in (Office.DocumentProperties)doc.BuiltInDocumentProperties)
                    {
                        list.Add(new WD.BuiltInDocumentProperty
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
                }) ?? new List<WD.BuiltInDocumentProperty>(),
                CustomDocumentProperties = Obj<List<WD.CustomDocumentProperty>>(() =>
                {
                    var list = new List<WD.CustomDocumentProperty>();
                    foreach (Office.DocumentProperty property in (Office.DocumentProperties)doc.CustomDocumentProperties)
                    {
                        list.Add(new WD.CustomDocumentProperty
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
                }) ?? new List<WD.CustomDocumentProperty>(),
                Bookmarks = Obj<List<WD.Bookmark>>(() =>
                {
                    var list = new List<WD.Bookmark>();
                    foreach (Bookmark bookmark in doc.Bookmarks)
                    {
                        list.Add(new WD.Bookmark
                        {
                            Name = Str(() => bookmark.Name),
                            RangeText = Str(() => bookmark.Range.Text),
                            Empty = Bool(() => bookmark.Empty),
                        });
                    }
                    return list;
                }) ?? new List<WD.Bookmark>(),
                Comments = Obj<List<WD.Comment>>(() =>
                {
                    var list = new List<WD.Comment>();
                    foreach (Comment comment in doc.Comments)
                    {
                        list.Add(new WD.Comment
                        {
                            Index = Long(() => comment.Index),
                            Author = Str(() => comment.Author),
                            Initial = Str(() => comment.Initial),
                            Date = DateStr(() => comment.Date),
                            RangeText = Str(() => comment.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Comment>(),
                ContentControls = Obj<List<WD.ContentControl>>(() =>
                {
                    var list = new List<WD.ContentControl>();
                    foreach (ContentControl control in doc.ContentControls)
                    {
                        list.Add(new WD.ContentControl
                        {
                            Title = Str(() => control.Title),
                            Tag = Str(() => control.Tag),
                            Type = Int(() => control.Type),
                            RangeText = Str(() => control.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.ContentControl>(),
                Endnotes = Obj<List<WD.Endnote>>(() =>
                {
                    var list = new List<WD.Endnote>();
                    foreach (Endnote endnote in doc.Endnotes)
                    {
                        list.Add(new WD.Endnote
                        {
                            Index = Long(() => endnote.Index),
                            RangeText = Str(() => endnote.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Endnote>(),
                Fields = Obj<List<WD.Field>>(() =>
                {
                    var list = new List<WD.Field>();
                    foreach (Field field in doc.Fields)
                    {
                        list.Add(new WD.Field
                        {
                            Index = Long(() => field.Index),
                            Type = Int(() => field.Type),
                            Locked = Bool(() => field.Locked),
                            Code = Str(() => field.Code.Text),
                            Result = Str(() => field.Result.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Field>(),
                Footnotes = Obj<List<WD.Footnote>>(() =>
                {
                    var list = new List<WD.Footnote>();
                    foreach (Footnote footnote in doc.Footnotes)
                    {
                        list.Add(new WD.Footnote
                        {
                            Index = Long(() => footnote.Index),
                            RangeText = Str(() => footnote.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Footnote>(),
                InlineShapes = Obj<List<WD.InlineShape>>(() =>
                {
                    var list = new List<WD.InlineShape>();
                    foreach (InlineShape shape in doc.InlineShapes)
                    {
                        list.Add(new WD.InlineShape
                        {
                            Image = null,
                            Type = Int(() => shape.Type),
                            Width = SingleStr(() => shape.Width),
                            Height = SingleStr(() => shape.Height),
                            AlternativeText = Str(() => shape.AlternativeText),
                        });
                    }
                    return list;
                }) ?? new List<WD.InlineShape>(),
                Shapes = Obj<List<WD.Shape>>(() =>
                {
                    var list = new List<WD.Shape>();
                    foreach (Shape shape in doc.Shapes)
                    {
                        list.Add(BuildShape(shape));
                    }
                    return list;
                }) ?? new List<WD.Shape>(),
                Paragraphs = Obj<List<WD.Paragraph>>(() =>
                {
                    var list = new List<WD.Paragraph>();
                    int index = 0;
                    int count = doc.Paragraphs.Count;
                    foreach (Paragraph paragraph in doc.Paragraphs)
                    {
                        index++;
                        if (index % 25 == 0) log?.Invoke($"  Paragraph {index}/{count}...");
                        list.Add(BuildParagraph(paragraph));
                    }
                    return list;
                }) ?? new List<WD.Paragraph>(),
                Sections = Obj<List<WD.Section>>(() =>
                {
                    var list = new List<WD.Section>();
                    foreach (Section section in doc.Sections)
                    {
                        list.Add(new WD.Section
                        {
                            PageSetup = Obj(() => BuildPageSetup(section.PageSetup)),
                            Index = Long(() => section.Index),
                        });
                    }
                    return list;
                }) ?? new List<WD.Section>(),
                Styles = Obj<List<WD.Style>>(() =>
                {
                    var list = new List<WD.Style>();
                    foreach (Style style in doc.Styles)
                    {
                        list.Add(new WD.Style
                        {
                            NameLocal = Str(() => style.NameLocal),
                            Type = Int(() => style.Type),
                            InUse = Bool(() => style.InUse),
                            BuiltIn = Bool(() => style.BuiltIn),
                            AutomaticallyUpdate = Bool(() => style.AutomaticallyUpdate),
                        });
                    }
                    return list;
                }) ?? new List<WD.Style>(),
                Tables = Obj<List<WD.Table>>(() =>
                {
                    var list = new List<WD.Table>();
                    foreach (Table table in doc.Tables)
                    {
                        list.Add(BuildTable(table));
                    }
                    return list;
                }) ?? new List<WD.Table>(),
                Variables = Obj<List<WD.Variable>>(() =>
                {
                    var list = new List<WD.Variable>();
                    foreach (Variable variable in doc.Variables)
                    {
                        list.Add(new WD.Variable
                        {
                            Name = Str(() => variable.Name),
                            Value = Str(() => variable.Value),
                        });
                    }
                    return list;
                }) ?? new List<WD.Variable>(),
                Name = Str(() => doc.Name),
                FullName = Str(() => doc.FullName),
                Path = Str(() => doc.Path),
                Saved = Bool(() => doc.Saved),
                SaveFormat = Int(() => doc.SaveFormat),
                ReadOnly = Bool(() => doc.ReadOnly),
                ProtectionType = Int(() => doc.ProtectionType),
                TrackRevisions = Bool(() => doc.TrackRevisions),
                VBASigned = Bool(() => doc.VBASigned),
            };

            return entity;
        }

        private static WD.Shape BuildShape(Shape shape)
        {
            var entity = new WD.Shape
            {
                Name = Str(() => shape.Name),
                Type = Int(() => shape.Type),
                Left = SingleStr(() => shape.Left),
                Top = SingleStr(() => shape.Top),
                Width = SingleStr(() => shape.Width),
                Height = SingleStr(() => shape.Height),
                Rotation = SingleStr(() => shape.Rotation),
                AlternativeText = Str(() => shape.AlternativeText),
                // Word's Shape does not expose HasTextFrame in the NetOffice wrappers;
                // the VBA writer reads it late-bound and writes null when the call fails.
                HasTextFrame = Bool(() => ((dynamic)shape).HasTextFrame),
            };

            if (entity.HasTextFrame == true)
            {
                entity.TextFrameText = Str(() => shape.TextFrame.TextRange.Text);
            }

            return entity;
        }

        private static WD.Paragraph BuildParagraph(Paragraph paragraph)
        {
            return new WD.Paragraph
            {
                Runs = Obj<List<WD.Run>>(() =>
                {
                    var list = new List<WD.Run>();
                    long index = 0;
                    foreach (Range character in paragraph.Range.Characters)
                    {
                        index++;
                        list.Add(new WD.Run
                        {
                            Font = Obj(() => BuildFont(character.Font)),
                            Index = index,
                            Text = Str(() => character.Text),
                        });

                        // Limit run count per paragraph to avoid excessively large output
                        if (index >= MaxRunsPerParagraph) break;
                    }
                    return list;
                }) ?? new List<WD.Run>(),
                Text = Str(() => paragraph.Range.Text),
                Style = Str(() => ((Style)paragraph.get_Style()).NameLocal),
                Alignment = Int(() => paragraph.Alignment),
                LeftIndent = SingleStr(() => paragraph.LeftIndent),
                RightIndent = SingleStr(() => paragraph.RightIndent),
                FirstLineIndent = SingleStr(() => paragraph.FirstLineIndent),
                SpaceBefore = SingleStr(() => paragraph.SpaceBefore),
                SpaceAfter = SingleStr(() => paragraph.SpaceAfter),
                LineSpacing = SingleStr(() => paragraph.LineSpacing),
                KeepWithNext = Bool(() => paragraph.KeepWithNext),
                PageBreakBefore = Bool(() => paragraph.PageBreakBefore),
            };
        }

        private static WD.Font BuildFont(Font font)
        {
            return new WD.Font
            {
                Name = Str(() => font.Name),
                Size = SingleStr(() => font.Size),
                Bold = Int(() => font.Bold),
                Italic = Int(() => font.Italic),
                Underline = Int(() => font.Underline),
                Color = Long(() => font.Color),
                ColorIndex = Int(() => font.ColorIndex),
                StrikeThrough = Int(() => font.StrikeThrough),
                Superscript = Int(() => font.Superscript),
                Subscript = Int(() => font.Subscript),
                Scaling = SingleStr(() => font.Scaling),
                Spacing = SingleStr(() => font.Spacing),
                Position = SingleStr(() => font.Position),
                Emboss = Int(() => font.Emboss),
                Shadow = Int(() => font.Shadow),
            };
        }

        private static WD.PageSetup BuildPageSetup(PageSetup pageSetup)
        {
            return new WD.PageSetup
            {
                PageWidth = SingleStr(() => pageSetup.PageWidth),
                PageHeight = SingleStr(() => pageSetup.PageHeight),
                Orientation = Int(() => pageSetup.Orientation),
                TopMargin = SingleStr(() => pageSetup.TopMargin),
                BottomMargin = SingleStr(() => pageSetup.BottomMargin),
                LeftMargin = SingleStr(() => pageSetup.LeftMargin),
                RightMargin = SingleStr(() => pageSetup.RightMargin),
                HeaderDistance = SingleStr(() => pageSetup.HeaderDistance),
                FooterDistance = SingleStr(() => pageSetup.FooterDistance),
                PaperSize = Int(() => pageSetup.PaperSize),
                SectionStart = Int(() => pageSetup.SectionStart),
                DifferentFirstPageHeaderFooter = Bool(() => pageSetup.DifferentFirstPageHeaderFooter),
            };
        }

        private static WD.Table BuildTable(Table table)
        {
            var entity = new WD.Table
            {
                Rows = new List<WD.TableRow>(),
                RowCount = Long(() => table.Rows.Count),
                ColumnCount = Long(() => table.Columns.Count),
            };

            try
            {
                foreach (Row row in table.Rows)
                {
                    var rowEntity = new WD.TableRow
                    {
                        Cells = new List<WD.TableCell>(),
                        Index = Long(() => row.Index),
                    };

                    foreach (Cell cell in row.Cells)
                    {
                        rowEntity.Cells.Add(new WD.TableCell
                        {
                            RowIndex = Long(() => cell.RowIndex),
                            ColumnIndex = Long(() => cell.ColumnIndex),
                            Text = Str(() => cell.Range.Text),
                        });
                    }

                    entity.Rows.Add(rowEntity);
                }
            }
            catch
            {
                // Same as VBA: a table whose rows cannot be enumerated keeps the rows read so far.
            }

            return entity;
        }
    }
}
