using System.Text;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;

namespace JBC.ExploreTheWorld.IntegrationTests.DL;

// Serialize/deserialize tests for the canonical "Save as JSON" entities
// (JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice / MsPowerPoint / MsExcel / MsWord).
public class MsOfficeJsonSerializer_Tests
{
    // ── PowerPoint ────────────────────────────────────────────────────────────────

    [Fact]
    public void Serialize_Presentation_RoundTripsAllValues()
    {
        var presentation = CreateSamplePresentation();

        string json = MsOfficeJsonSerializer.Serialize(presentation);
        var restored = MsOfficeJsonSerializer.Deserialize<PP.Presentation>(json);

        restored.Should().NotBeNull();
        restored!.Name.Should().Be("Sample.pptx");
        restored.HasTitleMaster.Should().Be(0);
        restored.PageSetup!.SlideWidth.Should().Be("960");
        restored.Slides.Should().HaveCount(1);

        var slide = restored.Slides![0];
        slide.SlideNumber.Should().Be(1);
        slide.Shapes.Should().HaveCount(2);
        slide.Shapes![0].TextFrame!.TextRange!.Text.Should().Be("Hello World");
        slide.Shapes[0].GroupItems.Should().BeNull();
        slide.Shapes[1].Type.Should().Be((int)PP.ShapeType_Enum.msoGroup);
        slide.Shapes[1].GroupItems.Should().HaveCount(1);
        slide.Shapes[1].GroupItems![0].Name.Should().Be("Grouped Picture");
        slide.Shapes[1].GroupItems[0].Image!.Base64.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void Serialize_Presentation_IsDeterministic()
    {
        var presentation = CreateSamplePresentation();

        string first = MsOfficeJsonSerializer.Serialize(presentation);
        string second = MsOfficeJsonSerializer.Serialize(
            MsOfficeJsonSerializer.Deserialize<PP.Presentation>(first)!);

        second.Should().Be(first);
    }

    [Fact]
    public void Serialize_Presentation_WritesEntitiesBeforeFields()
    {
        // The VBA writers emit sub-objects and lists first, then the scalar fields.
        // System.Text.Json must serialize the derived class (entities/lists) before
        // the base *_Fields class for the outputs to match.
        string json = MsOfficeJsonSerializer.Serialize(CreateSamplePresentation());

        // Root-level properties are indented with exactly two spaces.
        int slideMasterIndex = json.IndexOf("\r\n  \"SlideMaster\"", StringComparison.Ordinal);
        int slidesIndex = json.IndexOf("\r\n  \"Slides\"", StringComparison.Ordinal);
        int nameIndex = json.IndexOf("\r\n  \"Name\"", StringComparison.Ordinal);

        slideMasterIndex.Should().BeGreaterThanOrEqualTo(0);
        slideMasterIndex.Should().BeLessThan(slidesIndex);
        slidesIndex.Should().BeLessThan(nameIndex);
    }

    [Fact]
    public void Serialize_Presentation_UsesTwoSpaceIndentAndCrLf()
    {
        string json = MsOfficeJsonSerializer.Serialize(CreateSamplePresentation());

        json.Should().StartWith("{");
        json.Should().Contain("\r\n  \"SlideMaster\"");
        json.Should().NotContain("\n{", "the root brace must be the first character");
        json.Split('\n').Skip(1).Take(5).Should().OnlyContain(line => line.EndsWith("\r") || line.Length == 0);
    }

    [Fact]
    public void Serialize_NullEntities_AreWrittenExplicitly()
    {
        var presentation = new PP.Presentation();

        string json = MsOfficeJsonSerializer.Serialize(presentation);

        json.Should().Contain("\"SlideMaster\": null");
        json.Should().Contain("\"Coauthoring\": null");
        json.Should().Contain("\"Slides\": null");
    }

    [Fact]
    public void Deserialize_UndefinedMarkers_RoundTrip()
    {
        var shape = new PP.Shape
        {
            Name = MsOfficeUndefined.String,
            Left = MsOfficeUndefined.String,
            AutoShapeType = MsOfficeUndefined.Number,
            Id = MsOfficeUndefined.NumberLong,
        };

        string json = MsOfficeJsonSerializer.Serialize(shape);
        var restored = MsOfficeJsonSerializer.Deserialize<PP.Shape>(json)!;

        MsOfficeUndefined.IsUndefined(restored.Name).Should().BeTrue();
        MsOfficeUndefined.IsUndefined(restored.AutoShapeType).Should().BeTrue();
        MsOfficeUndefined.IsUndefined(restored.Id).Should().BeTrue();
        json.Should().Contain("\"Name\": \"**Undefined\"");
        json.Should().Contain("\"AutoShapeType\": -99");
    }

    [Fact]
    public void Deserialize_AcceptsTrailingCommas_FromLegacyVbaOutput()
    {
        // Files written by the pre-fix VBA writers contain trailing commas.
        string legacyJson = "{\r\n  \"Name\": \"Legacy.pptx\",\r\n}";

        var restored = MsOfficeJsonSerializer.Deserialize<PP.Presentation>(legacyJson);

        restored!.Name.Should().Be("Legacy.pptx");
    }

    // ── File output ──────────────────────────────────────────────────────────────

    [Fact]
    public void WriteToFile_WritesUtf8WithoutBomAndTrailingNewline()
    {
        string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".json");

        try
        {
            MsOfficeJsonSerializer.WriteToFile(CreateSamplePresentation(), path);

            byte[] bytes = File.ReadAllBytes(path);
            bytes.Take(3).Should().NotEqual(new byte[] { 0xEF, 0xBB, 0xBF }, "the file must not start with a BOM");
            bytes[0].Should().Be((byte)'{');
            Encoding.UTF8.GetString(bytes).Should().EndWith("}\r\n");

            var restored = MsOfficeJsonSerializer.ReadFromFile<PP.Presentation>(path);
            restored!.Name.Should().Be("Sample.pptx");
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ── Writer options (blob output) ─────────────────────────────────────────────

    [Fact]
    public void MsOfficeJsonWriterOptions_DefaultsToBase64()
    {
        var options = new MsOfficeJsonWriterOptions();

        options.BlobOutput.Should().Be(BlobOutput_Enum.Base64);
        options.BlobFolderPath.Should().BeNull();
    }

    [Fact]
    public void GetBlobFolderPath_DefaultsToFilesFolderBesideJson()
    {
        var options = new MsOfficeJsonWriterOptions();

        string folder = options.GetBlobFolderPath(@"C:\Exports\MyDeck.json");

        folder.Should().Be(@"C:\Exports\MyDeck_Files");
    }

    [Fact]
    public void GetBlobFolderPath_UsesExplicitFolderWhenSet()
    {
        var options = new MsOfficeJsonWriterOptions { BlobFolderPath = @"C:\Blobs" };

        options.GetBlobFolderPath(@"C:\Exports\MyDeck.json").Should().Be(@"C:\Blobs");
    }

    [Fact]
    public void ImageBlob_Base64AndFileName_AreMutuallyExclusiveByConvention()
    {
        var base64Blob = new ImageBlob { Extension = "png", Base64 = "AAAA", FileName = null };
        var fileBlob = new ImageBlob { Extension = "png", Base64 = null, FileName = "MyDeck_Files/5.png" };

        string base64Json = MsOfficeJsonSerializer.Serialize(base64Blob);
        string fileJson = MsOfficeJsonSerializer.Serialize(fileBlob);

        base64Json.Should().Contain("\"Base64\": \"AAAA\"").And.Contain("\"FileName\": null");
        fileJson.Should().Contain("\"Base64\": null").And.Contain("\"FileName\": \"MyDeck_Files/5.png\"");
    }

    // ── Excel ────────────────────────────────────────────────────────────────────

    [Fact]
    public void Serialize_Workbook_RoundTripsAllValues()
    {
        var workbook = new XL.Workbook
        {
            BuiltInDocumentProperties = new List<XL.BuiltInDocumentProperty>
            {
                new() { Name = "Title", Creator = MsOfficeUndefined.NumberLong, LinkToContent = 0, Type = 4, Value = "My Workbook" },
            },
            CustomDocumentProperties = new List<XL.CustomDocumentProperty>(),
            Sheets = new List<XL.Sheet>
            {
                new()
                {
                    UsedRange = new XL.UsedRange { Address = "$A$1:$B$2", RowCount = 2, ColumnCount = 2 },
                    Rows = new List<XL.Row>
                    {
                        new()
                        {
                            RowIndex = 1,
                            Cells = new List<XL.Cell>
                            {
                                new() { RowIndex = 1, ColumnIndex = 1, Address = "$A$1", Value = "Country", Formula = "Country" },
                                new() { RowIndex = 1, ColumnIndex = 2, Address = "$B$1", Value = "ISO2", Formula = "ISO2" },
                            },
                        },
                    },
                    Name = "Countries",
                    Index = 1,
                    Visible = -1,
                    Type = "Worksheet",
                },
            },
            Name = "Sample.xlsx",
            FullName = @"C:\Temp\Sample.xlsx",
            Path = @"C:\Temp",
            Saved = true,
            ReadOnly = false,
            HasVBProject = false,
            FileFormat = 51,
            CodeName = "ThisWorkbook",
        };

        string json = MsOfficeJsonSerializer.Serialize(workbook);
        var restored = MsOfficeJsonSerializer.Deserialize<XL.Workbook>(json)!;

        restored.Name.Should().Be("Sample.xlsx");
        restored.FileFormat.Should().Be(51);
        restored.Sheets.Should().HaveCount(1);
        restored.Sheets![0].UsedRange!.Address.Should().Be("$A$1:$B$2");
        restored.Sheets[0].Rows![0].Cells![1].Value.Should().Be("ISO2");

        // Entities/lists (derived class) serialize before the scalar fields (base class).
        json.IndexOf("\r\n  \"Sheets\"", StringComparison.Ordinal)
            .Should().BeLessThan(json.IndexOf("\r\n  \"Name\"", StringComparison.Ordinal));
    }

    // ── Word ─────────────────────────────────────────────────────────────────────

    [Fact]
    public void Serialize_Document_RoundTripsAllValues()
    {
        var document = new WD.Document
        {
            BuiltInDocumentProperties = new List<WD.BuiltInDocumentProperty>(),
            CustomDocumentProperties = new List<WD.CustomDocumentProperty>(),
            Bookmarks = new List<WD.Bookmark> { new() { Name = "Start", RangeText = "Hello", Empty = false } },
            Comments = new List<WD.Comment>(),
            ContentControls = new List<WD.ContentControl>(),
            Endnotes = new List<WD.Endnote>(),
            Fields = new List<WD.Field>(),
            Footnotes = new List<WD.Footnote>(),
            InlineShapes = new List<WD.InlineShape>
            {
                new() { Image = new ImageBlob { Extension = "png", Base64 = "QUJD" }, Type = 3, Width = "100", Height = "50", AlternativeText = "Logo" },
            },
            Shapes = new List<WD.Shape>(),
            Paragraphs = new List<WD.Paragraph>
            {
                new()
                {
                    Runs = new List<WD.Run>
                    {
                        new() { Font = new WD.Font { Name = "Calibri", Size = "11", Bold = -1 }, Index = 1, Text = "H" },
                    },
                    Text = "Hello",
                    Style = "Heading 1",
                    Alignment = 0,
                },
            },
            Sections = new List<WD.Section>
            {
                new() { PageSetup = new WD.PageSetup { PageWidth = "612", PageHeight = "792", Orientation = 0 }, Index = 1 },
            },
            Styles = new List<WD.Style>(),
            Tables = new List<WD.Table>
            {
                new()
                {
                    Rows = new List<WD.TableRow>
                    {
                        new()
                        {
                            Cells = new List<WD.TableCell> { new() { RowIndex = 1, ColumnIndex = 1, Text = "A1" } },
                            Index = 1,
                        },
                    },
                    RowCount = 1,
                    ColumnCount = 1,
                },
            },
            Variables = new List<WD.Variable>(),
            Name = "Sample.docx",
            FullName = @"C:\Temp\Sample.docx",
            Path = @"C:\Temp",
            Saved = true,
            SaveFormat = 16,
            ReadOnly = false,
            ProtectionType = -1,
            TrackRevisions = false,
            VBASigned = false,
        };

        string json = MsOfficeJsonSerializer.Serialize(document);
        var restored = MsOfficeJsonSerializer.Deserialize<WD.Document>(json)!;

        restored.Name.Should().Be("Sample.docx");
        restored.Paragraphs![0].Runs![0].Font!.Name.Should().Be("Calibri");
        restored.Tables![0].Rows![0].Cells![0].Text.Should().Be("A1");
        restored.InlineShapes![0].Image!.Base64.Should().Be("QUJD");
        restored.Sections![0].PageSetup!.PageWidth.Should().Be("612");

        // Entities/lists (derived class) serialize before the scalar fields (base class).
        json.IndexOf("\r\n  \"Paragraphs\"", StringComparison.Ordinal)
            .Should().BeLessThan(json.IndexOf("\r\n  \"Name\"", StringComparison.Ordinal));
        string paragraphJson = json.Substring(json.IndexOf("\"Paragraphs\"", StringComparison.Ordinal));
        paragraphJson.IndexOf("\"Runs\"", StringComparison.Ordinal)
            .Should().BeLessThan(paragraphJson.IndexOf("\"Text\"", StringComparison.Ordinal));
    }

    // ── Sample data ──────────────────────────────────────────────────────────────

    private static PP.Presentation CreateSamplePresentation()
    {
        return new PP.Presentation
        {
            SlideMaster = new PP.SlideMaster
            {
                SlideShowTransition = null,
                SlideMasterLayouts = new List<PP.SlideMasterLayout>
                {
                    new() { Shapes = new List<PP.Shape>(), Name = "Title Slide", Index = 1 },
                },
                Shapes = new List<PP.Shape>(),
                Name = "Office Theme",
            },
            Coauthoring = new PP.Coauthoring { FavorServerEditsDuringMerge = false, MergeMode = false, PendingUpdates = false },
            PageSetup = new PP.PageSetup
            {
                FirstSlideNumber = 1,
                NotesOrientation = 1,
                SlideHeight = "540",
                SlideOrientation = 2,
                SlideSize = 15,
                SlideWidth = "960",
            },
            BuiltInDocumentProperties = new List<PP.BuiltInDocumentProperty>
            {
                new() { Name = "Title", Creator = 0, LinkToContent = 0, Type = 4, Value = "Sample" },
            },
            CustomDocumentProperties = new List<PP.CustomDocumentProperty>(),
            SectionProperties = new List<PP.SectionProperty>
            {
                new() { Index = 1, Name = "Default Section", SectionID = "{1}", SlidesCount = 1 },
            },
            Slides = new List<PP.Slide>
            {
                new()
                {
                    Shapes = new List<PP.Shape>
                    {
                        new()
                        {
                            TextFrame = new PP.TextFrame
                            {
                                TextRange = new PP.TextRange { Font = null, Length = 11, Start = 1, Text = "Hello World" },
                            },
                            Tags = new List<PP.Tag>(),
                            Id = 2,
                            Name = "Title 1",
                            Type = (int)PP.ShapeType_Enum.msoPlaceholder,
                            Left = "66",
                            Top = "38",
                            Width = "828",
                            Height = "115",
                            ZOrderPosition = 1,
                            Visible = -1,
                        },
                        new()
                        {
                            Tags = new List<PP.Tag>(),
                            GroupItems = new List<PP.Shape>
                            {
                                new()
                                {
                                    Image = new ImageBlob { Extension = "png", Base64 = "iVBORw0KGgo=" },
                                    Tags = new List<PP.Tag>(),
                                    Id = 7,
                                    Name = "Grouped Picture",
                                    Type = (int)PP.ShapeType_Enum.msoPicture,
                                    ZOrderPosition = 1,
                                    Visible = -1,
                                },
                            },
                            Id = 6,
                            Name = "Group 5",
                            Type = (int)PP.ShapeType_Enum.msoGroup,
                            ZOrderPosition = 2,
                            Visible = -1,
                        },
                    },
                    Tags = new List<PP.Tag> { new() { Name = "STAGE", Value = "Draft" } },
                    Comments = null,
                    Hyperlinks = null,
                    Name = "Slide1",
                    CustomLayout_Name = "Title Slide",
                    SlideID = 256,
                    SlideIndex = 1,
                    SlideNumber = 1,
                },
            },
            Tags = new List<PP.Tag>(),
            Name = "Sample.pptx",
            Path = @"C:\Temp",
            FullName = @"C:\Temp\Sample.pptx",
            HasTitleMaster = 0,
            Saved = -1,
        };
    }
}
