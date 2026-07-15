using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using D = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using static JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters.OpenXmlJsonValues;
using Undefined = JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice.MsOfficeUndefined;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters
{
    // Builds the canonical JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord entity graph from a .docx file
    // and writes it with MsOfficeJsonSerializer. The schema matches
    // MSO_MsWord_JsonWriter.bas; VBA properties the OpenXML file format cannot provide
    // are written as "**Undefined" (strings), -99 (enums/numerics) or null (booleans).
    // Paragraph and run formatting is resolved through the style hierarchy (direct
    // properties, then the style basedOn chain, then docDefaults) so the output matches
    // the effective values the COM writers report.
    public static class MsWordJsonWriter
    {
        private const int MaxRunsPerParagraph = 500;

        // wdColorAutomatic
        private const long WdColorAutomatic = -16777216;

        // WdColorIndex palette (hex RRGGBB by index 1..16).
        private static readonly string[] WdColorIndexPalette =
        {
            "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF",
            "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0",
        };

        // Field instruction keyword to the WdFieldType value the COM writers report.
        private static readonly Dictionary<string, int> WdFieldTypes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["REF"] = 3, ["XE"] = 4, ["SET"] = 6, ["IF"] = 7, ["INDEX"] = 8, ["TC"] = 9,
            ["STYLEREF"] = 10, ["RD"] = 11, ["SEQ"] = 12, ["TOC"] = 13, ["INFO"] = 14,
            ["TITLE"] = 15, ["SUBJECT"] = 16, ["AUTHOR"] = 17, ["KEYWORDS"] = 18, ["COMMENTS"] = 19,
            ["LASTSAVEDBY"] = 20, ["CREATEDATE"] = 21, ["SAVEDATE"] = 22, ["PRINTDATE"] = 23,
            ["REVNUM"] = 24, ["EDITTIME"] = 25, ["NUMPAGES"] = 26, ["NUMWORDS"] = 27, ["NUMCHARS"] = 28,
            ["FILENAME"] = 29, ["TEMPLATE"] = 30, ["DATE"] = 31, ["TIME"] = 32, ["PAGE"] = 33,
            ["QUOTE"] = 35, ["INCLUDE"] = 36, ["PAGEREF"] = 37, ["ASK"] = 38, ["FILLIN"] = 39,
            ["NEXT"] = 41, ["NEXTIF"] = 42, ["SKIPIF"] = 43, ["MERGEREC"] = 44, ["PRINT"] = 48,
            ["EQ"] = 49, ["GOTOBUTTON"] = 50, ["MACROBUTTON"] = 51, ["AUTONUM"] = 54, ["LINK"] = 56,
            ["SYMBOL"] = 57, ["EMBED"] = 58, ["MERGEFIELD"] = 59, ["USERNAME"] = 60, ["USERINITIALS"] = 61,
            ["USERADDRESS"] = 62, ["BARCODE"] = 63, ["DOCVARIABLE"] = 64, ["SECTION"] = 65,
            ["SECTIONPAGES"] = 66, ["INCLUDEPICTURE"] = 67, ["INCLUDETEXT"] = 68, ["FILESIZE"] = 69,
            ["FORMTEXT"] = 70, ["FORMCHECKBOX"] = 71, ["NOTEREF"] = 72, ["TOA"] = 73, ["TA"] = 74,
            ["MERGESEQ"] = 75, ["PRIVATE"] = 77, ["DATABASE"] = 78, ["AUTOTEXT"] = 79, ["COMPARE"] = 80,
            ["ADDIN"] = 81, ["FORMDROPDOWN"] = 83, ["ADVANCE"] = 84, ["DOCPROPERTY"] = 85,
            ["HYPERLINK"] = 88, ["AUTOTEXTLIST"] = 89, ["LISTNUM"] = 90, ["BIDIOUTLINE"] = 92,
            ["ADDRESSBLOCK"] = 93, ["GREETINGLINE"] = 94,
        };

        public static void WriteDocumentFileToJson(string sourcePath, string outputJsonPath,
            MsOfficeJsonWriterOptions? options = null, Action<string>? log = null)
        {
            options ??= new MsOfficeJsonWriterOptions();

            log?.Invoke($"Writing (OpenXml): {outputJsonPath}");

            using var document = WordprocessingDocument.Open(sourcePath, false);

            WD.Document entity = BuildDocument(document, sourcePath, outputJsonPath, options, log);

            MsOfficeJsonSerializer.WriteToFile(entity, outputJsonPath);

            log?.Invoke($"Done: {outputJsonPath}");
        }

        private static WD.Document BuildDocument(WordprocessingDocument document, string sourcePath,
            string outputJsonPath, MsOfficeJsonWriterOptions options, Action<string>? log)
        {
            var mainPart = document.MainDocumentPart
                ?? throw new InvalidOperationException("Cannot read main document part.");
            var body = mainPart.Document?.Body;

            var entity = new WD.Document
            {
                BuiltInDocumentProperties = OpenXmlDocumentProperties.BuildBuiltIn(document.PackageProperties,
                    document.ExtendedFilePropertiesPart,
                    (name, value, type) => new WD.BuiltInDocumentProperty
                    {
                        Name = name,
                        Creator = Undefined.NumberLong,
                        LinkSource = null,
                        LinkToContent = 0,
                        Type = type,
                        Value = value,
                    }),
                CustomDocumentProperties = OpenXmlDocumentProperties.BuildCustom(document.CustomFilePropertiesPart,
                    (name, value, type) => new WD.CustomDocumentProperty
                    {
                        Name = name,
                        Creator = Undefined.NumberLong,
                        LinkSource = null,
                        LinkToContent = 0,
                        Type = type,
                        Value = value,
                    }),
                Bookmarks = BuildBookmarks(body),
                Comments = BuildComments(mainPart),
                ContentControls = BuildContentControls(body),
                Endnotes = BuildEndnotes(mainPart),
                Fields = BuildFields(body),
                Footnotes = BuildFootnotes(mainPart),
                InlineShapes = BuildInlineShapes(mainPart, outputJsonPath, options),
                Shapes = BuildShapes(mainPart),
                Paragraphs = BuildParagraphs(mainPart, log),
                Sections = BuildSections(body),
                Styles = BuildStyles(mainPart),
                Tables = BuildTables(body),
                Variables = BuildVariables(mainPart),
                Name = Path.GetFileName(sourcePath),
                FullName = sourcePath,
                Path = Path.GetDirectoryName(sourcePath) ?? string.Empty,
                Saved = null,
                SaveFormat = Path.GetExtension(sourcePath).ToLowerInvariant() switch
                {
                    ".docx" => 12, // wdFormatXMLDocument
                    ".docm" => 13, // wdFormatXMLDocumentMacroEnabled
                    ".dotx" => 14, // wdFormatXMLTemplate
                    ".dotm" => 15, // wdFormatXMLTemplateMacroEnabled
                    _ => Undefined.Number,
                },
                ReadOnly = null,
                ProtectionType = BuildProtectionType(mainPart),
                TrackRevisions = mainPart.DocumentSettingsPart?.Settings?
                    .ChildElements.Any(e => e.LocalName == "trackChanges"),
                VBASigned = null,
            };

            return entity;
        }

        private static int BuildProtectionType(MainDocumentPart mainPart)
        {
            var protection = mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<W.DocumentProtection>();

            // Protection that is present but not enforced reads as wdNoProtection, like COM.
            if (protection == null || protection.Enforcement?.Value != true) return -1; // wdNoProtection

            return protection.Edit?.Value switch
            {
                null => -1,
                W.DocumentProtectionValues edit when edit == W.DocumentProtectionValues.ReadOnly => 3, // wdAllowOnlyReading
                W.DocumentProtectionValues edit when edit == W.DocumentProtectionValues.Comments => 1, // wdAllowOnlyComments
                W.DocumentProtectionValues edit when edit == W.DocumentProtectionValues.TrackedChanges => 0, // wdAllowOnlyRevisions
                W.DocumentProtectionValues edit when edit == W.DocumentProtectionValues.Forms => 2, // wdAllowOnlyFormFields
                _ => -1,
            };
        }

        private static List<WD.Bookmark> BuildBookmarks(W.Body? body)
        {
            var list = new List<WD.Bookmark>();
            if (body == null) return list;

            // Single pass in document order so each bookmark's range text can be captured
            // between its bookmarkStart and bookmarkEnd markers.
            var openBookmarks = new Dictionary<string, StringBuilder>(StringComparer.Ordinal);
            var textByBookmark = new List<(string Name, StringBuilder Text)>();

            foreach (var element in body.Descendants())
            {
                switch (element)
                {
                    case W.BookmarkStart bookmarkStart:
                        string? name = bookmarkStart.Name?.Value;
                        string? id = bookmarkStart.Id?.Value;
                        if (name == null || id == null || name.StartsWith("_", StringComparison.Ordinal)) continue;

                        var builder = new StringBuilder();
                        openBookmarks[id] = builder;
                        textByBookmark.Add((name, builder));
                        break;

                    case W.BookmarkEnd bookmarkEnd:
                        if (bookmarkEnd.Id?.Value is string endId) openBookmarks.Remove(endId);
                        break;

                    case W.Text text:
                        foreach (var open in openBookmarks.Values) open.Append(text.Text);
                        break;
                }
            }

            foreach (var (name, text) in textByBookmark)
            {
                list.Add(new WD.Bookmark
                {
                    Name = name,
                    RangeText = text.ToString(),
                    Empty = text.Length == 0,
                });
            }

            return list;
        }

        private static List<WD.Comment> BuildComments(MainDocumentPart mainPart)
        {
            var list = new List<WD.Comment>();

            var comments = mainPart.WordprocessingCommentsPart?.Comments;
            if (comments == null) return list;

            long index = 0;
            foreach (var comment in comments.Elements<W.Comment>())
            {
                index++;
                list.Add(new WD.Comment
                {
                    Index = index,
                    Author = comment.Author?.Value ?? string.Empty,
                    Initial = comment.Initials?.Value ?? string.Empty,
                    Date = comment.Date?.Value is DateTime date
                        ? date.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture)
                        : null,
                    RangeText = comment.InnerText,
                });
            }

            return list;
        }

        private static List<WD.ContentControl> BuildContentControls(W.Body? body)
        {
            var list = new List<WD.ContentControl>();
            if (body == null) return list;

            foreach (var sdt in body.Descendants<W.SdtElement>())
            {
                var properties = sdt.SdtProperties;
                list.Add(new WD.ContentControl
                {
                    Title = properties?.GetFirstChild<W.SdtAlias>()?.Val?.Value ?? string.Empty,
                    Tag = properties?.GetFirstChild<W.Tag>()?.Val?.Value ?? string.Empty,
                    Type = ContentControlTypeFrom(properties),
                    RangeText = sdt.InnerText,
                });
            }

            return list;
        }

        private static int ContentControlTypeFrom(W.SdtProperties? properties)
        {
            // WdContentControlType values; rich text (0) is the default control type.
            if (properties == null) return 0;

            if (properties.GetFirstChild<W.SdtContentText>() != null) return 1; // wdContentControlText
            if (properties.GetFirstChild<W.SdtContentPicture>() != null) return 2; // wdContentControlPicture
            if (properties.GetFirstChild<W.SdtContentComboBox>() != null) return 3; // wdContentControlComboBox
            if (properties.GetFirstChild<W.SdtContentDropDownList>() != null) return 4; // wdContentControlDropdownList
            if (properties.GetFirstChild<W.SdtContentDocPartObject>() != null
                || properties.GetFirstChild<W.SdtContentDocPartList>() != null) return 5; // wdContentControlBuildingBlockGallery
            if (properties.GetFirstChild<W.SdtContentDate>() != null) return 6; // wdContentControlDate
            if (properties.GetFirstChild<W.SdtContentGroup>() != null) return 7; // wdContentControlGroup
            if (properties.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>() != null) return 8; // wdContentControlCheckBox
            if (properties.GetFirstChild<DocumentFormat.OpenXml.Office2013.Word.SdtRepeatedSection>() != null) return 9; // wdContentControlRepeatingSection

            return 0; // wdContentControlRichText
        }

        private static List<WD.Endnote> BuildEndnotes(MainDocumentPart mainPart)
        {
            var list = new List<WD.Endnote>();

            var endnotes = mainPart.EndnotesPart?.Endnotes;
            if (endnotes == null) return list;

            long index = 0;
            foreach (var endnote in endnotes.Elements<W.Endnote>())
            {
                // Skip the built-in separator/continuation notes (negative ids).
                if ((endnote.Id?.Value ?? 0) < 1) continue;

                index++;
                list.Add(new WD.Endnote
                {
                    Index = index,
                    RangeText = endnote.InnerText,
                });
            }

            return list;
        }

        private static List<WD.Field> BuildFields(W.Body? body)
        {
            var list = new List<WD.Field>();
            if (body == null) return list;

            long index = 0;

            // Complex-field state (fldChar begin/separate/end around instrText and result runs).
            // Nested fields only contribute their text to the outermost field, which matches
            // how the COM writers flatten Document.Fields content closely enough.
            int depth = 0;
            bool inResult = false;
            bool locked = false;
            StringBuilder? code = null;
            StringBuilder? result = null;

            foreach (var element in body.Descendants())
            {
                switch (element)
                {
                    case W.SimpleField simpleField:
                        index++;
                        string instruction = simpleField.Instruction?.Value ?? string.Empty;
                        list.Add(new WD.Field
                        {
                            Index = index,
                            Type = FieldTypeFromCode(instruction),
                            Locked = simpleField.FieldLock?.Value ?? false,
                            Code = instruction,
                            Result = simpleField.InnerText,
                        });
                        break;

                    case W.FieldChar fieldChar:
                        var charType = fieldChar.FieldCharType?.Value;
                        if (charType == W.FieldCharValues.Begin)
                        {
                            depth++;
                            if (depth == 1)
                            {
                                code = new StringBuilder();
                                result = new StringBuilder();
                                inResult = false;
                                locked = fieldChar.FieldLock?.Value ?? false;
                            }
                        }
                        else if (charType == W.FieldCharValues.Separate)
                        {
                            if (depth == 1) inResult = true;
                        }
                        else if (charType == W.FieldCharValues.End)
                        {
                            if (depth == 1 && code != null)
                            {
                                index++;
                                list.Add(new WD.Field
                                {
                                    Index = index,
                                    Type = FieldTypeFromCode(code.ToString()),
                                    Locked = locked,
                                    Code = code.ToString(),
                                    Result = result?.ToString() ?? string.Empty,
                                });
                                code = null;
                                result = null;
                            }
                            if (depth > 0) depth--;
                        }
                        break;

                    case W.FieldCode fieldCode:
                        if (depth == 1 && !inResult) code?.Append(fieldCode.Text);
                        break;

                    case W.Text text:
                        if (depth == 1 && inResult) result?.Append(text.Text);
                        break;
                }
            }

            return list;
        }

        private static int FieldTypeFromCode(string? code)
        {
            string trimmed = code?.Trim() ?? string.Empty;
            if (trimmed.Length == 0) return Undefined.Number;

            if (trimmed.StartsWith("=", StringComparison.Ordinal)) return 34; // wdFieldExpression

            string keyword = new string(trimmed.TakeWhile(char.IsLetter).ToArray());

            return WdFieldTypes.TryGetValue(keyword, out int type) ? type : Undefined.Number;
        }

        private static List<WD.Footnote> BuildFootnotes(MainDocumentPart mainPart)
        {
            var list = new List<WD.Footnote>();

            var footnotes = mainPart.FootnotesPart?.Footnotes;
            if (footnotes == null) return list;

            long index = 0;
            foreach (var footnote in footnotes.Elements<W.Footnote>())
            {
                if ((footnote.Id?.Value ?? 0) < 1) continue;

                index++;
                list.Add(new WD.Footnote
                {
                    Index = index,
                    RangeText = footnote.InnerText,
                });
            }

            return list;
        }

        private static List<WD.InlineShape> BuildInlineShapes(MainDocumentPart mainPart, string outputJsonPath,
            MsOfficeJsonWriterOptions options)
        {
            var list = new List<WD.InlineShape>();
            var body = mainPart.Document?.Body;
            if (body == null) return list;

            long imageIndex = 0;
            foreach (var inline in body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>())
            {
                imageIndex++;

                var extent = inline.Extent;
                var docProperties = inline.DocProperties;

                list.Add(new WD.InlineShape
                {
                    Image = BuildImage(inline, mainPart, imageIndex, outputJsonPath, options),
                    Type = 3, // wdInlineShapePicture
                    Width = PointsFromEmu(extent?.Cx?.Value),
                    Height = PointsFromEmu(extent?.Cy?.Value),
                    AlternativeText = docProperties?.Description?.Value ?? string.Empty,
                });
            }

            return list;
        }

        private static ImageBlob? BuildImage(DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline,
            MainDocumentPart mainPart, long imageIndex, string outputJsonPath, MsOfficeJsonWriterOptions options)
        {
            try
            {
                string? embed = inline.Descendants<D.Blip>().FirstOrDefault()?.Embed?.Value;
                if (embed == null) return null;

                if (mainPart.GetPartById(embed) is not ImagePart imagePart) return null;

                using var stream = imagePart.GetStream();
                using var memory = new MemoryStream();
                stream.CopyTo(memory);
                byte[] bytes = memory.ToArray();

                string extension = ExtensionFromContentType(imagePart.ContentType);
                var blob = new ImageBlob { Extension = extension };

                if (options.BlobOutput == BlobOutput_Enum.Base64)
                {
                    blob.Base64 = Convert.ToBase64String(bytes);
                }
                else
                {
                    string blobFolder = options.GetBlobFolderPath(outputJsonPath);
                    Directory.CreateDirectory(blobFolder);

                    string fileName = $"{imageIndex}.{extension}";
                    File.WriteAllBytes(Path.Combine(blobFolder, fileName), bytes);

                    blob.FileName = Path.GetFileName(blobFolder) + "/" + fileName;
                }

                return blob;
            }
            catch
            {
                return null;
            }
        }

        private static List<WD.Shape> BuildShapes(MainDocumentPart mainPart)
        {
            var list = new List<WD.Shape>();
            var body = mainPart.Document?.Body;
            if (body == null) return list;

            foreach (var anchor in body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>())
            {
                var docProperties = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>();
                var extent = anchor.Extent;

                string? textFrameText = anchor.Descendants<W.TextBoxContent>().FirstOrDefault()?.InnerText;

                list.Add(new WD.Shape
                {
                    Name = docProperties?.Name?.Value ?? Undefined.String,
                    Type = ShapeTypeFrom(anchor, textFrameText),
                    Left = PointsFromEmu(anchor.HorizontalPosition?.PositionOffset?.Text is string x
                        ? long.Parse(x, CultureInfo.InvariantCulture)
                        : null),
                    Top = PointsFromEmu(anchor.VerticalPosition?.PositionOffset?.Text is string y
                        ? long.Parse(y, CultureInfo.InvariantCulture)
                        : null),
                    Width = PointsFromEmu(extent?.Cx?.Value),
                    Height = PointsFromEmu(extent?.Cy?.Value),
                    Rotation = DegreesFrom60000ths(anchor.Descendants<D.Transform2D>().FirstOrDefault()?.Rotation?.Value),
                    AlternativeText = docProperties?.Description?.Value ?? string.Empty,
                    HasTextFrame = textFrameText != null,
                    TextFrameText = textFrameText,
                });
            }

            return list;
        }

        private static int ShapeTypeFrom(DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor anchor, string? textFrameText)
        {
            // MsoShapeType values derived from the anchored drawing's graphic data URI.
            string? uri = anchor.Descendants<D.Graphic>().FirstOrDefault()?.GraphicData?.Uri?.Value;

            return uri switch
            {
                "http://schemas.openxmlformats.org/drawingml/2006/picture" => 13, // msoPicture
                "http://schemas.openxmlformats.org/drawingml/2006/chart" => 3, // msoChart
                "http://schemas.openxmlformats.org/drawingml/2006/diagram" => 24, // msoSmartArt
                "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" => 6, // msoGroup
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" =>
                    textFrameText != null
                        && anchor.Descendants<D.PresetGeometry>().FirstOrDefault()?.Preset?.Value == D.ShapeTypeValues.Rectangle
                        ? 17 // msoTextBox
                        : 1, // msoAutoShape
                _ => Undefined.Number,
            };
        }

        private static List<WD.Paragraph> BuildParagraphs(MainDocumentPart mainPart, Action<string>? log)
        {
            var list = new List<WD.Paragraph>();
            var body = mainPart.Document?.Body;
            if (body == null) return list;

            var styles = BuildStyleMap(mainPart);
            var docDefaults = mainPart.StyleDefinitionsPart?.Styles?.DocDefaults;
            var fontScheme = mainPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            string? defaultParagraphStyleId = styles.Values
                .FirstOrDefault(s => s.Type?.Value == W.StyleValues.Paragraph && s.Default?.Value == true)?
                .StyleId?.Value;

            var paragraphs = body.Descendants<W.Paragraph>().ToList();

            int index = 0;
            foreach (var paragraph in paragraphs)
            {
                index++;
                if (index % 25 == 0) log?.Invoke($"  Paragraph {index}/{paragraphs.Count}...");

                list.Add(BuildParagraph(paragraph, styles, docDefaults, defaultParagraphStyleId, fontScheme));
            }

            return list;
        }

        private static Dictionary<string, W.Style> BuildStyleMap(MainDocumentPart mainPart)
        {
            var map = new Dictionary<string, W.Style>(StringComparer.Ordinal);

            var styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) return map;

            foreach (var style in styles.Elements<W.Style>())
            {
                if (style.StyleId?.Value is string styleId) map[styleId] = style;
            }

            return map;
        }

        // Effective paragraph property: direct pPr, then the paragraph style basedOn chain,
        // then the document defaults.
        private static TElement? ResolveParagraphElement<TElement>(W.Paragraph paragraph, string? styleId,
            Dictionary<string, W.Style> styles, W.DocDefaults? docDefaults) where TElement : OpenXmlElement
        {
            var direct = paragraph.ParagraphProperties?.GetFirstChild<TElement>();
            if (direct != null) return direct;

            string? currentId = styleId;
            for (int depth = 0; depth < 16 && currentId != null && styles.TryGetValue(currentId, out var style); depth++)
            {
                var fromStyle = style.StyleParagraphProperties?.GetFirstChild<TElement>();
                if (fromStyle != null) return fromStyle;

                currentId = style.BasedOn?.Val?.Value;
            }

            return docDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.GetFirstChild<TElement>();
        }

        // Effective run property: direct rPr, then the run character style chain, then the
        // paragraph style chain, then the document defaults.
        private static TElement? ResolveRunElement<TElement>(W.Run run, string? paragraphStyleId,
            Dictionary<string, W.Style> styles, W.DocDefaults? docDefaults) where TElement : OpenXmlElement
        {
            var direct = run.RunProperties?.GetFirstChild<TElement>();
            if (direct != null) return direct;

            foreach (string? startId in new[] { run.RunProperties?.RunStyle?.Val?.Value, paragraphStyleId })
            {
                string? currentId = startId;
                for (int depth = 0; depth < 16 && currentId != null && styles.TryGetValue(currentId, out var style); depth++)
                {
                    var fromStyle = style.StyleRunProperties?.GetFirstChild<TElement>();
                    if (fromStyle != null) return fromStyle;

                    currentId = style.BasedOn?.Val?.Value;
                }
            }

            return docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<TElement>();
        }

        private static bool OnOff(W.OnOffType? element)
        {
            return element != null && (element.Val?.Value ?? true);
        }

        private static WD.Paragraph BuildParagraph(W.Paragraph paragraph, Dictionary<string, W.Style> styles,
            W.DocDefaults? docDefaults, string? defaultParagraphStyleId, D.FontScheme? fontScheme)
        {
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? defaultParagraphStyleId;
            string styleName = styleId != null && styles.TryGetValue(styleId, out var style)
                ? style.StyleName?.Val?.Value ?? styleId
                : styleId ?? "Normal";

            var indentation = ResolveParagraphElement<W.Indentation>(paragraph, styleId, styles, docDefaults);
            var spacing = ResolveParagraphElement<W.SpacingBetweenLines>(paragraph, styleId, styles, docDefaults);

            double? firstLineIndent = ParseTwips(indentation?.FirstLine?.Value);
            if (firstLineIndent == null && ParseTwips(indentation?.Hanging?.Value) is double hanging)
            {
                firstLineIndent = -hanging;
            }

            var entity = new WD.Paragraph
            {
                Runs = new List<WD.Run>(),
                Text = paragraph.InnerText,
                Style = styleName,
                Alignment = ResolveParagraphElement<W.Justification>(paragraph, styleId, styles, docDefaults)?.Val?.Value switch
                {
                    null => 0, // wdAlignParagraphLeft
                    W.JustificationValues justification when justification == W.JustificationValues.Left => 0,
                    W.JustificationValues justification when justification == W.JustificationValues.Start => 0,
                    W.JustificationValues justification when justification == W.JustificationValues.Center => 1,
                    W.JustificationValues justification when justification == W.JustificationValues.Right => 2,
                    W.JustificationValues justification when justification == W.JustificationValues.End => 2,
                    W.JustificationValues justification when justification == W.JustificationValues.Both => 3,
                    W.JustificationValues justification when justification == W.JustificationValues.Distribute => 4, // wdAlignParagraphDistribute
                    _ => Undefined.Number,
                },
                LeftIndent = PointsFromTwips(ParseTwips(indentation?.Left?.Value ?? indentation?.Start?.Value) ?? 0),
                RightIndent = PointsFromTwips(ParseTwips(indentation?.Right?.Value ?? indentation?.End?.Value) ?? 0),
                FirstLineIndent = PointsFromTwips(firstLineIndent ?? 0),
                SpaceBefore = PointsFromTwips(ParseTwips(spacing?.Before?.Value) ?? 0),
                SpaceAfter = PointsFromTwips(ParseTwips(spacing?.After?.Value) ?? 0),
                LineSpacing = PointsFromTwips(ParseTwips(spacing?.Line?.Value)) is string line && line != Undefined.String
                    ? line
                    : Undefined.String,
                KeepWithNext = OnOff(ResolveParagraphElement<W.KeepNext>(paragraph, styleId, styles, docDefaults)),
                PageBreakBefore = OnOff(ResolveParagraphElement<W.PageBreakBefore>(paragraph, styleId, styles, docDefaults)),
            };

            long runIndex = 0;
            foreach (var run in paragraph.Elements<W.Run>())
            {
                runIndex++;
                entity.Runs.Add(new WD.Run
                {
                    Font = BuildFont(run, styleId, styles, docDefaults, fontScheme),
                    Index = runIndex,
                    Text = run.InnerText,
                });

                // Limit run count per paragraph to avoid excessively large output
                if (runIndex >= MaxRunsPerParagraph) break;
            }

            return entity;
        }

        private static double? ParseTwips(string? value)
        {
            if (string.IsNullOrEmpty(value)) return null;

            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips)
                ? twips
                : null;
        }

        private static WD.Font BuildFont(W.Run run, string? paragraphStyleId, Dictionary<string, W.Style> styles,
            W.DocDefaults? docDefaults, D.FontScheme? fontScheme)
        {
            var runFonts = ResolveRunElement<W.RunFonts>(run, paragraphStyleId, styles, docDefaults);
            var color = ResolveRunElement<W.Color>(run, paragraphStyleId, styles, docDefaults);
            var verticalAlignment = ResolveRunElement<W.VerticalTextAlignment>(run, paragraphStyleId, styles, docDefaults)?.Val?.Value;

            return new WD.Font
            {
                Name = runFonts?.Ascii?.Value
                    ?? ThemeFontName(fontScheme, runFonts?.AsciiTheme?.Value)
                    ?? Undefined.String,
                Size = PointsFromHalfPoints(ParseTwips(
                    ResolveRunElement<W.FontSize>(run, paragraphStyleId, styles, docDefaults)?.Val?.Value)),
                Bold = TriState(OnOff(ResolveRunElement<W.Bold>(run, paragraphStyleId, styles, docDefaults))),
                Italic = TriState(OnOff(ResolveRunElement<W.Italic>(run, paragraphStyleId, styles, docDefaults))),
                Underline = WdUnderlineFrom(ResolveRunElement<W.Underline>(run, paragraphStyleId, styles, docDefaults)),
                Color = WdColorFromValue(color?.Val?.Value),
                ColorIndex = WdColorIndexFromValue(color?.Val?.Value),
                StrikeThrough = TriState(OnOff(ResolveRunElement<W.Strike>(run, paragraphStyleId, styles, docDefaults))),
                Superscript = TriState(verticalAlignment == W.VerticalPositionValues.Superscript),
                Subscript = TriState(verticalAlignment == W.VerticalPositionValues.Subscript),
                Scaling = ResolveRunElement<W.CharacterScale>(run, paragraphStyleId, styles, docDefaults)?.Val?.Value is long scaling
                    ? scaling.ToString(CultureInfo.InvariantCulture)
                    : "100",
                Spacing = PointsFromTwips((double?)ResolveRunElement<W.Spacing>(run, paragraphStyleId, styles, docDefaults)?.Val?.Value) is string spacing
                    && spacing != Undefined.String ? spacing : "0",
                Position = PointsFromHalfPoints(ParseTwips(
                    ResolveRunElement<W.Position>(run, paragraphStyleId, styles, docDefaults)?.Val?.Value)) is string position
                    && position != Undefined.String ? position : "0",
                Emboss = TriState(OnOff(ResolveRunElement<W.Emboss>(run, paragraphStyleId, styles, docDefaults))),
                Shadow = TriState(OnOff(ResolveRunElement<W.Shadow>(run, paragraphStyleId, styles, docDefaults))),
            };
        }

        private static string? ThemeFontName(D.FontScheme? fontScheme, W.ThemeFontValues? themeFont)
        {
            if (fontScheme == null || themeFont == null) return null;

            var value = themeFont.Value;

            if (value == W.ThemeFontValues.MajorAscii || value == W.ThemeFontValues.MajorHighAnsi)
            {
                return fontScheme.MajorFont?.LatinFont?.Typeface?.Value;
            }

            if (value == W.ThemeFontValues.MinorAscii || value == W.ThemeFontValues.MinorHighAnsi)
            {
                return fontScheme.MinorFont?.LatinFont?.Typeface?.Value;
            }

            return null;
        }

        private static int WdUnderlineFrom(W.Underline? underline)
        {
            return underline?.Val?.Value switch
            {
                null => 0, // wdUnderlineNone
                W.UnderlineValues value when value == W.UnderlineValues.None => 0,
                W.UnderlineValues value when value == W.UnderlineValues.Single => 1,
                W.UnderlineValues value when value == W.UnderlineValues.Words => 2,
                W.UnderlineValues value when value == W.UnderlineValues.Double => 3,
                W.UnderlineValues value when value == W.UnderlineValues.Dotted => 4,
                W.UnderlineValues value when value == W.UnderlineValues.Thick => 6,
                W.UnderlineValues value when value == W.UnderlineValues.Dash => 7,
                W.UnderlineValues value when value == W.UnderlineValues.DotDash => 9,
                W.UnderlineValues value when value == W.UnderlineValues.DotDotDash => 10,
                W.UnderlineValues value when value == W.UnderlineValues.Wave => 11,
                W.UnderlineValues value when value == W.UnderlineValues.DottedHeavy => 20,
                W.UnderlineValues value when value == W.UnderlineValues.DashedHeavy => 23,
                W.UnderlineValues value when value == W.UnderlineValues.DashDotHeavy => 25,
                W.UnderlineValues value when value == W.UnderlineValues.DashDotDotHeavy => 26,
                W.UnderlineValues value when value == W.UnderlineValues.WavyHeavy => 27,
                W.UnderlineValues value when value == W.UnderlineValues.DashLong => 39,
                W.UnderlineValues value when value == W.UnderlineValues.WavyDouble => 43,
                W.UnderlineValues value when value == W.UnderlineValues.DashLongHeavy => 55,
                _ => Undefined.Number,
            };
        }

        private static long? WdColorFromValue(string? value)
        {
            if (value == null || value == "auto") return WdColorAutomatic;

            return OleColorFromHex(value);
        }

        private static int WdColorIndexFromValue(string? value)
        {
            if (value == null || value == "auto") return 0; // wdAuto

            for (int index = 0; index < WdColorIndexPalette.Length; index++)
            {
                if (string.Equals(WdColorIndexPalette[index], value, StringComparison.OrdinalIgnoreCase))
                {
                    return index + 1;
                }
            }

            return Undefined.Number;
        }

        private static List<WD.Section> BuildSections(W.Body? body)
        {
            var list = new List<WD.Section>();
            if (body == null) return list;

            var sectionProperties = body.Descendants<W.SectionProperties>().ToList();

            long index = 0;
            foreach (var section in sectionProperties)
            {
                index++;

                var pageSize = section.GetFirstChild<W.PageSize>();
                var pageMargin = section.GetFirstChild<W.PageMargin>();

                list.Add(new WD.Section
                {
                    PageSetup = new WD.PageSetup
                    {
                        PageWidth = PointsFromTwips((double?)pageSize?.Width?.Value),
                        PageHeight = PointsFromTwips((double?)pageSize?.Height?.Value),
                        Orientation = pageSize?.Orient?.Value == W.PageOrientationValues.Landscape ? 1 : 0,
                        TopMargin = PointsFromTwips((double?)pageMargin?.Top?.Value),
                        BottomMargin = PointsFromTwips((double?)pageMargin?.Bottom?.Value),
                        LeftMargin = PointsFromTwips((double?)pageMargin?.Left?.Value),
                        RightMargin = PointsFromTwips((double?)pageMargin?.Right?.Value),
                        HeaderDistance = PointsFromTwips((double?)pageMargin?.Header?.Value),
                        FooterDistance = PointsFromTwips((double?)pageMargin?.Footer?.Value),
                        PaperSize = WdPaperSizeFrom(pageSize),
                        SectionStart = section.GetFirstChild<W.SectionType>()?.Val?.Value switch
                        {
                            null => 2, // wdSectionNewPage (the OOXML default)
                            W.SectionMarkValues mark when mark == W.SectionMarkValues.Continuous => 0,
                            W.SectionMarkValues mark when mark == W.SectionMarkValues.NextColumn => 1,
                            W.SectionMarkValues mark when mark == W.SectionMarkValues.NextPage => 2,
                            W.SectionMarkValues mark when mark == W.SectionMarkValues.EvenPage => 3,
                            W.SectionMarkValues mark when mark == W.SectionMarkValues.OddPage => 4,
                            _ => Undefined.Number,
                        },
                        DifferentFirstPageHeaderFooter = section.GetFirstChild<W.TitlePage>() != null,
                    },
                    Index = index,
                });
            }

            return list;
        }

        private static int WdPaperSizeFrom(W.PageSize? pageSize)
        {
            // The pgSz code attribute holds the Windows DMPAPER value the printer reported.
            switch (pageSize?.Code?.Value)
            {
                case 1: return 2; // wdPaperLetter
                case 2: return 3; // wdPaperLetterSmall
                case 3: return 1; // wdPaper11x17
                case 5: return 4; // wdPaperLegal
                case 7: return 5; // wdPaperExecutive
                case 8: return 6; // wdPaperA3
                case 9: return 7; // wdPaperA4
                case 10: return 8; // wdPaperA4Small
                case 11: return 9; // wdPaperA5
                case 12: return 10; // wdPaperB4
                case 13: return 11; // wdPaperB5
            }

            // No code attribute: match the page dimensions (twips) against the common sizes.
            if (pageSize?.Width?.Value is uint width && pageSize.Height?.Value is uint height)
            {
                uint shortSide = Math.Min(width, height);
                uint longSide = Math.Max(width, height);

                if (shortSide == 12240 && longSide == 15840) return 2; // wdPaperLetter
                if (shortSide == 11906 && longSide == 16838) return 7; // wdPaperA4
                if (shortSide == 12240 && longSide == 20160) return 4; // wdPaperLegal
                if (shortSide == 16838 && longSide == 23811) return 6; // wdPaperA3
                if (shortSide == 8391 && longSide == 11906) return 9; // wdPaperA5
            }

            return Undefined.Number;
        }

        private static List<WD.Style> BuildStyles(MainDocumentPart mainPart)
        {
            var list = new List<WD.Style>();

            var styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) return list;

            var body = mainPart.Document?.Body;
            var usedStyleIds = new HashSet<string>(StringComparer.Ordinal);
            if (body != null)
            {
                foreach (var element in body.Descendants())
                {
                    string? styleId = element switch
                    {
                        W.ParagraphStyleId paragraphStyle => paragraphStyle.Val?.Value,
                        W.RunStyle runStyle => runStyle.Val?.Value,
                        W.TableStyle tableStyle => tableStyle.Val?.Value,
                        _ => null,
                    };

                    if (styleId != null) usedStyleIds.Add(styleId);
                }
            }

            foreach (var style in styles.Elements<W.Style>())
            {
                list.Add(new WD.Style
                {
                    NameLocal = style.StyleName?.Val?.Value ?? style.StyleId?.Value ?? string.Empty,
                    Type = style.Type?.Value switch
                    {
                        null => Undefined.Number,
                        W.StyleValues type when type == W.StyleValues.Paragraph => 1, // wdStyleTypeParagraph
                        W.StyleValues type when type == W.StyleValues.Character => 2, // wdStyleTypeCharacter
                        W.StyleValues type when type == W.StyleValues.Table => 3, // wdStyleTypeTable
                        W.StyleValues type when type == W.StyleValues.Numbering => 4, // wdStyleTypeList
                        _ => Undefined.Number,
                    },
                    InUse = style.Default?.Value == true
                        || (style.StyleId?.Value is string styleId && usedStyleIds.Contains(styleId)),
                    BuiltIn = style.CustomStyle?.Value != true,
                    AutomaticallyUpdate = style.AutoRedefine != null,
                });
            }

            return list;
        }

        private static List<WD.Table> BuildTables(W.Body? body)
        {
            var list = new List<WD.Table>();
            if (body == null) return list;

            foreach (var table in body.Elements<W.Table>())
            {
                var rows = table.Elements<W.TableRow>().ToList();

                var entity = new WD.Table
                {
                    Rows = new List<WD.TableRow>(),
                    RowCount = rows.Count,
                    ColumnCount = rows.Count == 0 ? 0 : rows.Max(r => r.Elements<W.TableCell>().Count()),
                };

                long rowIndex = 0;
                foreach (var row in rows)
                {
                    rowIndex++;
                    var rowEntity = new WD.TableRow
                    {
                        Cells = new List<WD.TableCell>(),
                        Index = rowIndex,
                    };

                    long columnIndex = 0;
                    foreach (var cell in row.Elements<W.TableCell>())
                    {
                        columnIndex++;
                        rowEntity.Cells.Add(new WD.TableCell
                        {
                            RowIndex = rowIndex,
                            ColumnIndex = columnIndex,
                            Text = cell.InnerText,
                        });
                    }

                    entity.Rows.Add(rowEntity);
                }

                list.Add(entity);
            }

            return list;
        }

        private static List<WD.Variable> BuildVariables(MainDocumentPart mainPart)
        {
            var list = new List<WD.Variable>();

            var docVariables = mainPart.DocumentSettingsPart?.Settings?
                .GetFirstChild<W.DocumentVariables>();
            if (docVariables == null) return list;

            foreach (var variable in docVariables.Elements<W.DocumentVariable>())
            {
                list.Add(new WD.Variable
                {
                    Name = variable.Name?.Value ?? string.Empty,
                    Value = variable.Val?.Value ?? string.Empty,
                });
            }

            return list;
        }
    }
}
