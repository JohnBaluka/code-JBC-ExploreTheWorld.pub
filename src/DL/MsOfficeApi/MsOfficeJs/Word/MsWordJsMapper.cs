using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word
{
    // Maps the Office.js-shaped rows to the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord
    // entities and back. VBA properties the Office.js object model cannot provide are set
    // to "**Undefined" (strings), -99 (enums/numerics) or null (booleans/whole objects).
    public static class MsWordJsMapper
    {
        public static WD.Document ToDocument(WordDocumentJs_Row source, string? fileName)
        {
            var document = new WD.Document
            {
                BuiltInDocumentProperties = ToBuiltInDocumentProperties(source.Properties),
                CustomDocumentProperties = new List<WD.CustomDocumentProperty>(),
                Bookmarks = new List<WD.Bookmark>(),
                Comments = new List<WD.Comment>(),
                ContentControls = new List<WD.ContentControl>(),
                Endnotes = new List<WD.Endnote>(),
                Fields = new List<WD.Field>(),
                Footnotes = new List<WD.Footnote>(),
                InlineShapes = new List<WD.InlineShape>(),
                Shapes = new List<WD.Shape>(),
                Paragraphs = new List<WD.Paragraph>(),
                Sections = new List<WD.Section>(),
                Styles = new List<WD.Style>(),
                Tables = new List<WD.Table>(),
                Variables = new List<WD.Variable>(),
                Name = string.IsNullOrEmpty(fileName) ? Undefined.String : fileName,
                FullName = Undefined.String,
                Path = Undefined.String,
                Saved = null,
                SaveFormat = Undefined.Number,
                ReadOnly = null,
                ProtectionType = Undefined.Number,
                TrackRevisions = null,
                VBASigned = null,
            };

            foreach (var contentControl in source.ContentControls)
            {
                document.ContentControls.Add(new WD.ContentControl
                {
                    Title = contentControl.Title ?? string.Empty,
                    Tag = contentControl.Tag ?? string.Empty,
                    Type = Undefined.Number,
                    RangeText = contentControl.Text ?? string.Empty,
                });
            }

            foreach (var picture in source.InlinePictures)
            {
                document.InlineShapes.Add(new WD.InlineShape
                {
                    Image = ToImageBlob(picture.Base64),
                    Type = 3, // wdInlineShapePicture
                    Width = PointsString(picture.Width),
                    Height = PointsString(picture.Height),
                    AlternativeText = picture.AltTextDescription ?? string.Empty,
                });
            }

            foreach (var paragraph in source.Paragraphs)
            {
                document.Paragraphs.Add(ToParagraph(paragraph));
            }

            foreach (var table in source.Tables)
            {
                document.Tables.Add(ToTable(table));
            }

            return document;
        }

        private static List<WD.BuiltInDocumentProperty> ToBuiltInDocumentProperties(WordDocumentPropertiesJs_Row? properties)
        {
            WD.BuiltInDocumentProperty Property(string name, string? value, int type = 4)
            {
                return new WD.BuiltInDocumentProperty
                {
                    Name = name,
                    Creator = Undefined.NumberLong,
                    LinkSource = null,
                    LinkToContent = 0,
                    Type = type,
                    Value = value,
                };
            }

            return new List<WD.BuiltInDocumentProperty>
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

        private static ImageBlob? ToImageBlob(string? base64)
        {
            if (string.IsNullOrEmpty(base64)) return null;

            string data = base64!;
            string extension = "png";

            // getBase64ImageSrc may return a data URI ("data:image/png;base64,...").
            if (data.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                int comma = data.IndexOf(',');
                if (comma > 0)
                {
                    string header = data.Substring(5, comma - 5);
                    int semicolon = header.IndexOf(';');
                    if (semicolon > 0)
                    {
                        extension = header.Substring(0, semicolon) switch
                        {
                            "image/png" => "png",
                            "image/jpeg" => "jpg",
                            "image/gif" => "gif",
                            "image/bmp" => "bmp",
                            _ => "png",
                        };
                    }

                    data = data.Substring(comma + 1);
                }
            }

            return new ImageBlob
            {
                Extension = extension,
                Base64 = data,
                FileName = null,
            };
        }

        private static WD.Paragraph ToParagraph(WordParagraphJs_Row source)
        {
            return new WD.Paragraph
            {
                Runs = new List<WD.Run>(),
                Text = source.Text ?? string.Empty,
                Style = source.Style ?? Undefined.String,
                Alignment = source.Alignment?.ToLowerInvariant() switch
                {
                    "left" => 0, // wdAlignParagraphLeft
                    "centered" => 1, // wdAlignParagraphCenter
                    "right" => 2, // wdAlignParagraphRight
                    "justified" => 3, // wdAlignParagraphJustify
                    _ => Undefined.Number,
                },
                LeftIndent = PointsString(source.LeftIndent),
                RightIndent = PointsString(source.RightIndent),
                FirstLineIndent = PointsString(source.FirstLineIndent),
                SpaceBefore = PointsString(source.SpaceBefore),
                SpaceAfter = PointsString(source.SpaceAfter),
                LineSpacing = PointsString(source.LineSpacing),
                KeepWithNext = null,
                PageBreakBefore = null,
            };
        }

        private static WD.Table ToTable(WordTableJs_Row source)
        {
            var table = new WD.Table
            {
                Rows = new List<WD.TableRow>(),
                RowCount = source.RowCount,
                ColumnCount = source.Values is { Count: > 0 } ? source.Values[0].Count : 0,
            };

            long rowIndex = 0;
            foreach (var values in source.Values ?? new List<List<string?>>())
            {
                rowIndex++;
                var row = new WD.TableRow
                {
                    Cells = new List<WD.TableCell>(),
                    Index = rowIndex,
                };

                long columnIndex = 0;
                foreach (string? value in values)
                {
                    columnIndex++;
                    row.Cells.Add(new WD.TableCell
                    {
                        RowIndex = rowIndex,
                        ColumnIndex = columnIndex,
                        Text = value ?? string.Empty,
                    });
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private static string PointsString(double? points)
        {
            if (!points.HasValue) return Undefined.String;

            return ((float)points.Value).ToString(CultureInfo.InvariantCulture);
        }

        public static WordDocumentJs_Row FromDocument(WD.Document document)
        {
            var row = new WordDocumentJs_Row
            {
                Paragraphs = new List<WordParagraphJs_Row>(),
                Tables = new List<WordTableJs_Row>(),
                ContentControls = new List<WordContentControlJs_Row>(),
                InlinePictures = new List<WordInlinePictureJs_Row>(),
            };

            foreach (var paragraph in document.Paragraphs ?? new List<WD.Paragraph>())
            {
                row.Paragraphs.Add(new WordParagraphJs_Row
                {
                    Text = paragraph.Text,
                    Style = paragraph.Style,
                    Alignment = paragraph.Alignment switch
                    {
                        0 => "Left",
                        1 => "Centered",
                        2 => "Right",
                        3 => "Justified",
                        _ => null,
                    },
                    LeftIndent = PointsFromString(paragraph.LeftIndent),
                    RightIndent = PointsFromString(paragraph.RightIndent),
                    FirstLineIndent = PointsFromString(paragraph.FirstLineIndent),
                    SpaceBefore = PointsFromString(paragraph.SpaceBefore),
                    SpaceAfter = PointsFromString(paragraph.SpaceAfter),
                    LineSpacing = PointsFromString(paragraph.LineSpacing),
                });
            }

            foreach (var table in document.Tables ?? new List<WD.Table>())
            {
                row.Tables.Add(new WordTableJs_Row
                {
                    RowCount = table.RowCount ?? 0,
                    Values = (table.Rows ?? new List<WD.TableRow>())
                        .Select(r => (r.Cells ?? new List<WD.TableCell>()).Select(c => c.Text).ToList())
                        .ToList(),
                });
            }

            foreach (var contentControl in document.ContentControls ?? new List<WD.ContentControl>())
            {
                row.ContentControls.Add(new WordContentControlJs_Row
                {
                    Title = contentControl.Title,
                    Tag = contentControl.Tag,
                    Text = contentControl.RangeText,
                });
            }

            foreach (var inlineShape in document.InlineShapes ?? new List<WD.InlineShape>())
            {
                row.InlinePictures.Add(new WordInlinePictureJs_Row
                {
                    Base64 = inlineShape.Image?.Base64,
                    Width = PointsFromString(inlineShape.Width),
                    Height = PointsFromString(inlineShape.Height),
                    AltTextDescription = inlineShape.AlternativeText,
                });
            }

            return row;
        }

        private static double? PointsFromString(string? points)
        {
            if (string.IsNullOrEmpty(points) || MsOfficeUndefined.IsUndefined(points)) return null;

            return double.TryParse(points, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : null;
        }
    }
}
