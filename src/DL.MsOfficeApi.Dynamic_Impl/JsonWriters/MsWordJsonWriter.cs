using System;
using System.Collections;
using System.Collections.Generic;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters.ComTryGet;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entity graph from a running
    // Word document COM object (late-bound or Interop PIA) and writes it with
    // MsOfficeJsonSerializer. The output matches MSO_MsWord_JsonWriter.bas.
    // Like the VBA writer, the InlineShape Image property is always null on this path
    // (the OpenXML writer provides the image bytes); the writer options are accepted
    // for signature parity.
    public static class MsWordJsonWriter
    {
        private const int MaxRunsPerParagraph = 500;

        // Accepts any Word Document COM object (late-bound or Interop PIA) —
        // used by the Dynamic repos and the Save-As-JSON helpers.
        public static void WriteDocumentComToJsonFile(object documentComObject, string outputFilePath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            log?.Invoke($"Writing JSON: {outputFilePath}");

            dynamic com = documentComObject;

            WD.Document entity = BuildDocument(com, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputFilePath);

            log?.Invoke($"Done: {outputFilePath}");
        }

        private static WD.Document BuildDocument(dynamic doc, Action<string>? log)
        {
            var entity = new WD.Document
            {
                BuiltInDocumentProperties = Obj<List<WD.BuiltInDocumentProperty>>(() =>
                {
                    var list = new List<WD.BuiltInDocumentProperty>();
                    foreach (dynamic property in (IEnumerable)doc.BuiltInDocumentProperties)
                    {
                        dynamic px = property;
                        list.Add(new WD.BuiltInDocumentProperty
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
                }) ?? new List<WD.BuiltInDocumentProperty>(),
                CustomDocumentProperties = Obj<List<WD.CustomDocumentProperty>>(() =>
                {
                    var list = new List<WD.CustomDocumentProperty>();
                    foreach (dynamic property in (IEnumerable)doc.CustomDocumentProperties)
                    {
                        dynamic px = property;
                        list.Add(new WD.CustomDocumentProperty
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
                }) ?? new List<WD.CustomDocumentProperty>(),
                Bookmarks = Obj<List<WD.Bookmark>>(() =>
                {
                    var list = new List<WD.Bookmark>();
                    foreach (dynamic bookmark in doc.Bookmarks)
                    {
                        dynamic bx = bookmark;
                        list.Add(new WD.Bookmark
                        {
                            Name = Str(() => bx.Name),
                            RangeText = Str(() => bx.Range.Text),
                            Empty = Bool(() => bx.Empty),
                        });
                    }
                    return list;
                }) ?? new List<WD.Bookmark>(),
                Comments = Obj<List<WD.Comment>>(() =>
                {
                    var list = new List<WD.Comment>();
                    foreach (dynamic comment in doc.Comments)
                    {
                        dynamic cx = comment;
                        list.Add(new WD.Comment
                        {
                            Index = Long(() => cx.Index),
                            Author = Str(() => cx.Author),
                            Initial = Str(() => cx.Initial),
                            Date = DateStr(() => cx.Date),
                            RangeText = Str(() => cx.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Comment>(),
                ContentControls = Obj<List<WD.ContentControl>>(() =>
                {
                    var list = new List<WD.ContentControl>();
                    foreach (dynamic control in doc.ContentControls)
                    {
                        dynamic ccx = control;
                        list.Add(new WD.ContentControl
                        {
                            Title = Str(() => ccx.Title),
                            Tag = Str(() => ccx.Tag),
                            Type = Int(() => ccx.Type),
                            RangeText = Str(() => ccx.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.ContentControl>(),
                Endnotes = Obj<List<WD.Endnote>>(() =>
                {
                    var list = new List<WD.Endnote>();
                    foreach (dynamic endnote in doc.Endnotes)
                    {
                        dynamic ex = endnote;
                        list.Add(new WD.Endnote
                        {
                            Index = Long(() => ex.Index),
                            RangeText = Str(() => ex.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Endnote>(),
                Fields = Obj<List<WD.Field>>(() =>
                {
                    var list = new List<WD.Field>();
                    foreach (dynamic field in doc.Fields)
                    {
                        dynamic fx = field;
                        list.Add(new WD.Field
                        {
                            Index = Long(() => fx.Index),
                            Type = Int(() => fx.Type),
                            Locked = Bool(() => fx.Locked),
                            Code = Str(() => fx.Code.Text),
                            Result = Str(() => fx.Result.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Field>(),
                Footnotes = Obj<List<WD.Footnote>>(() =>
                {
                    var list = new List<WD.Footnote>();
                    foreach (dynamic footnote in doc.Footnotes)
                    {
                        dynamic fx = footnote;
                        list.Add(new WD.Footnote
                        {
                            Index = Long(() => fx.Index),
                            RangeText = Str(() => fx.Range.Text),
                        });
                    }
                    return list;
                }) ?? new List<WD.Footnote>(),
                InlineShapes = Obj<List<WD.InlineShape>>(() =>
                {
                    var list = new List<WD.InlineShape>();
                    foreach (dynamic shape in doc.InlineShapes)
                    {
                        dynamic sx = shape;
                        list.Add(new WD.InlineShape
                        {
                            Image = null,
                            Type = Int(() => sx.Type),
                            Width = SingleStr(() => sx.Width),
                            Height = SingleStr(() => sx.Height),
                            AlternativeText = Str(() => sx.AlternativeText),
                        });
                    }
                    return list;
                }) ?? new List<WD.InlineShape>(),
                Shapes = Obj<List<WD.Shape>>(() =>
                {
                    var list = new List<WD.Shape>();
                    foreach (dynamic shape in doc.Shapes)
                    {
                        list.Add(BuildShape(shape));
                    }
                    return list;
                }) ?? new List<WD.Shape>(),
                Paragraphs = Obj<List<WD.Paragraph>>(() =>
                {
                    var list = new List<WD.Paragraph>();
                    int index = 0;
                    int count = Convert.ToInt32(doc.Paragraphs.Count);
                    foreach (dynamic paragraph in doc.Paragraphs)
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
                    foreach (dynamic section in doc.Sections)
                    {
                        dynamic sx = section;
                        list.Add(new WD.Section
                        {
                            PageSetup = Obj(() => BuildPageSetup(sx.PageSetup)),
                            Index = Long(() => sx.Index),
                        });
                    }
                    return list;
                }) ?? new List<WD.Section>(),
                Styles = Obj<List<WD.Style>>(() =>
                {
                    var list = new List<WD.Style>();
                    foreach (dynamic style in doc.Styles)
                    {
                        dynamic stx = style;
                        list.Add(new WD.Style
                        {
                            NameLocal = Str(() => stx.NameLocal),
                            Type = Int(() => stx.Type),
                            InUse = Bool(() => stx.InUse),
                            BuiltIn = Bool(() => stx.BuiltIn),
                            AutomaticallyUpdate = Bool(() => stx.AutomaticallyUpdate),
                        });
                    }
                    return list;
                }) ?? new List<WD.Style>(),
                Tables = Obj<List<WD.Table>>(() =>
                {
                    var list = new List<WD.Table>();
                    foreach (dynamic table in doc.Tables)
                    {
                        list.Add(BuildTable(table));
                    }
                    return list;
                }) ?? new List<WD.Table>(),
                Variables = Obj<List<WD.Variable>>(() =>
                {
                    var list = new List<WD.Variable>();
                    foreach (dynamic variable in doc.Variables)
                    {
                        dynamic vx = variable;
                        list.Add(new WD.Variable
                        {
                            Name = Str(() => vx.Name),
                            Value = Str(() => vx.Value),
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

        private static WD.Shape BuildShape(dynamic shape)
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
                HasTextFrame = Bool(() => shape.HasTextFrame),
            };

            if (entity.HasTextFrame == true)
            {
                entity.TextFrameText = Str(() => shape.TextFrame.TextRange.Text);
            }

            return entity;
        }

        private static WD.Paragraph BuildParagraph(dynamic paragraph)
        {
            return new WD.Paragraph
            {
                Runs = Obj<List<WD.Run>>(() =>
                {
                    var list = new List<WD.Run>();
                    long index = 0;
                    foreach (dynamic character in paragraph.Range.Characters)
                    {
                        index++;
                        dynamic cx = character;
                        list.Add(new WD.Run
                        {
                            Font = Obj(() => BuildFont(cx.Font)),
                            Index = index,
                            Text = Str(() => cx.Text),
                        });

                        // Limit run count per paragraph to avoid excessively large output
                        if (index >= MaxRunsPerParagraph) break;
                    }
                    return list;
                }) ?? new List<WD.Run>(),
                Text = Str(() => paragraph.Range.Text),
                Style = Str(() => paragraph.Style.NameLocal),
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

        private static WD.Font BuildFont(dynamic font)
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

        private static WD.PageSetup BuildPageSetup(dynamic pageSetup)
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

        private static WD.Table BuildTable(dynamic table)
        {
            var entity = new WD.Table
            {
                Rows = new List<WD.TableRow>(),
                RowCount = Long(() => table.Rows.Count),
                ColumnCount = Long(() => table.Columns.Count),
            };

            try
            {
                foreach (dynamic row in table.Rows)
                {
                    dynamic rx = row;
                    var rowEntity = new WD.TableRow
                    {
                        Cells = new List<WD.TableCell>(),
                        Index = Long(() => rx.Index),
                    };

                    foreach (dynamic cell in rx.Cells)
                    {
                        dynamic cx = cell;
                        rowEntity.Cells.Add(new WD.TableCell
                        {
                            RowIndex = Long(() => cx.RowIndex),
                            ColumnIndex = Long(() => cx.ColumnIndex),
                            Text = Str(() => cx.Range.Text),
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
