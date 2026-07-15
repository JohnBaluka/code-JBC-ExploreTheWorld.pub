using System.Text.Json.Nodes;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;

namespace JBC.ExploreTheWorld.WinFormAppTests.JsonWriters;

// The VBA "Direct" writer is the gold standard for the canonical Save-As-JSON schema.
// These tests exercise the WinFormApp watcher write methods (NetOffice / Dynamic /
// OpenXml) against the same presentation and pin the behaviors the COM writers were
// aligned to: G7 single formatting, per-property nulls instead of 0-defaults, and
// MsoTriState values for the Has* presentation properties. The NetOffice and Dynamic
// tests return early when PowerPoint is not installed.
public class MsPowerPointJsonWriterComparison_Tests : IClassFixture<PowerPointJsonOutputs_Fixture>
{
    private readonly PowerPointJsonOutputs_Fixture _fixture;

    public MsPowerPointJsonWriterComparison_Tests(PowerPointJsonOutputs_Fixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void OpenXmlOutput_DeserializesIntoCanonicalEntities()
    {
        var presentation = _fixture.OpenXmlPresentation;

        presentation.Should().NotBeNull();
        presentation!.Slides.Should().HaveCount(2, "a title slide plus one data slide");
    }

    [Fact]
    public void NetOfficeAndDynamicOutputs_AreIdentical()
    {
        if (!_fixture.PowerPointAvailable) return;

        // Both writers read the same closed file through a hidden PowerPoint instance,
        // so their canonical output must match byte-for-byte at the JSON level. The file
        // is opened Untitled, so PowerPoint assigns a session-scoped "PresentationN"
        // name — Name/FullName are the only fields allowed to differ.
        var netOffice = JsonNode.Parse(File.ReadAllText(_fixture.NetOfficeJsonPath))!.AsObject();
        var dynamicJson = JsonNode.Parse(File.ReadAllText(_fixture.DynamicJsonPath))!.AsObject();

        foreach (var output in new[] { netOffice, dynamicJson })
        {
            output["Name"] = null;
            output["FullName"] = null;
        }

        JsonNode.DeepEquals(netOffice, dynamicJson).Should().BeTrue(
            "the NetOffice (typed) and Dynamic (late-bound) writers read the same COM object model");
    }

    [Fact]
    public void NetOfficeOutput_MatchesGoldStandardSemantics()
    {
        if (!_fixture.PowerPointAvailable) return;

        var presentation = _fixture.NetOfficePresentation;
        presentation.Should().NotBeNull();

        // MsoTriState values, not bool-derived 1/0 (VBA writes msoTrue as -1).
        presentation!.HasHandoutMaster.Should().Be(-1);
        presentation.HasNotesMaster.Should().Be(-1);
        presentation.HasVBProject.Should().Be(0);

        // VBA CStr(Single) formatting: at most 7 significant digits, no long tails.
        var geometry = presentation.Slides!
            .SelectMany(slide => slide.Shapes ?? new List<PP.Shape>())
            .SelectMany(shape => new[] { shape.Left, shape.Top, shape.Width, shape.Height })
            .Where(value => value != null)
            .ToList();
        geometry.Should().NotBeEmpty();
        geometry.Should().OnlyContain(value => CountSignificantDigits(value!) <= 7);

        // A property that cannot be converted (Shape.Vertices returns an array) is null,
        // never the .NET type name.
        presentation.Slides!
            .SelectMany(slide => slide.Shapes ?? new List<PP.Shape>())
            .Should().OnlyContain(shape => shape.Vertices != "System.Single[,]");

        // The PlaySettings getter throws for non-media shapes, so the object is null.
        presentation.Slides!
            .SelectMany(slide => slide.Shapes ?? new List<PP.Shape>())
            .Where(shape => shape.AnimationSettings != null)
            .Should().OnlyContain(shape => shape.AnimationSettings!.PlaySettings == null);
    }

    [Fact]
    public void OpenXmlOutput_AgreesWithNetOfficeOnCoreStructure()
    {
        if (!_fixture.PowerPointAvailable) return;

        var netOffice = _fixture.NetOfficePresentation!;
        var openXml = _fixture.OpenXmlPresentation!;

        openXml.Slides.Should().HaveCount(netOffice.Slides!.Count);

        for (int slideIndex = 0; slideIndex < netOffice.Slides.Count; slideIndex++)
        {
            var netOfficeSlide = netOffice.Slides[slideIndex];
            var openXmlSlide = openXml.Slides![slideIndex];

            openXmlSlide.Shapes.Should().HaveCount(netOfficeSlide.Shapes!.Count,
                $"slide {slideIndex + 1} has the same shapes in the file and the COM model");

            for (int shapeIndex = 0; shapeIndex < netOfficeSlide.Shapes.Count; shapeIndex++)
            {
                var netOfficeShape = netOfficeSlide.Shapes[shapeIndex];
                var openXmlShape = openXmlSlide.Shapes![shapeIndex];

                openXmlShape.Name.Should().Be(netOfficeShape.Name);
                openXmlShape.TextFrame?.TextRange?.Text.Should().Be(netOfficeShape.TextFrame?.TextRange?.Text);

                // Where the OpenXml writer resolves geometry it must agree with the
                // effective values COM reports (Undefined markers are allowed while
                // OpenXml cannot derive a value).
                AssertGeometryAgrees(openXmlShape.Left, netOfficeShape.Left, "Left");
                AssertGeometryAgrees(openXmlShape.Top, netOfficeShape.Top, "Top");
                AssertGeometryAgrees(openXmlShape.Width, netOfficeShape.Width, "Width");
                AssertGeometryAgrees(openXmlShape.Height, netOfficeShape.Height, "Height");
            }
        }
    }

    private static void AssertGeometryAgrees(string? openXmlValue, string? comValue, string propertyName)
    {
        if (openXmlValue == null || comValue == null) return;
        if (MsOfficeUndefined.IsUndefined(openXmlValue)) return;

        double openXml = double.Parse(openXmlValue, System.Globalization.CultureInfo.InvariantCulture);
        double com = double.Parse(comValue, System.Globalization.CultureInfo.InvariantCulture);

        openXml.Should().BeApproximately(com, 0.05, $"{propertyName} must match the COM value");
    }

    private static int CountSignificantDigits(string value)
    {
        var digits = value.Split('E', 'e')[0].TrimStart('-').Replace(".", "").TrimStart('0');
        return digits.TrimEnd('0').Length is 0 ? 1 : digits.Length;
    }
}
