using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;

namespace JBC.ExploreTheWorld.OpenXmlLibTests.Managers;

public class MsPowerPoint_OpenXml__Repo_Tests
{
    private static readonly List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS"),
        new MsOfficeCountry_Row("Brazil",    "BR", "BRA")
    };

    private readonly MsPowerPoint_OpenXml__Repo _sut = new();

    [Fact]
    public async Task ExportAsync_CreatesPresentation_WithTitleAndOneSlidePerCountry()
    {
        var pptxPath = TestFiles.NewTempPath(".pptx");
        try
        {
            await _sut.ExportAsync(Countries, pptxPath, _ => { });

            File.Exists(pptxPath).Should().BeTrue();

            using var presentation = PresentationDocument.Open(pptxPath, false);
            var slideParts = presentation.PresentationPart!.SlideParts.ToList();

            slideParts.Should().HaveCount(Countries.Count + 1, "a title slide plus one slide per country");

            var allText = string.Concat(slideParts.Select(sp => sp.Slide!.InnerText));
            allText.Should().Contain("Explore the World").And.Contain("Australia").And.Contain("Brazil");
        }
        finally
        {
            TestFiles.Delete(pptxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_WithFlagPng_EmbedsPicturePerFlaggedRow()
    {
        var countriesWithFlags = new List<MsOfficeCountry_Row>
        {
            new("Australia", "AU", "AUS", TestFiles.TinyPng),
            new("Brazil",    "BR", "BRA", TestFiles.TinyPng),
            new("Nowhere",   "XX", "XXX") // no flag — must not add a picture
        };
        var pptxPath = TestFiles.NewTempPath(".pptx");
        try
        {
            await _sut.ExportAsync(countriesWithFlags, pptxPath, _ => { });

            using var presentation = PresentationDocument.Open(pptxPath, false);
            var slideParts = presentation.PresentationPart!.SlideParts.ToList();

            slideParts.Sum(sp => sp.ImageParts.Count()).Should().Be(2);
            slideParts
                .Sum(sp => sp.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Count())
                .Should().Be(2);
        }
        finally
        {
            TestFiles.Delete(pptxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_ReportsProgress_ThroughLog()
    {
        var pptxPath = TestFiles.NewTempPath(".pptx");
        var logLines = new List<string>();
        try
        {
            await _sut.ExportAsync(Countries, pptxPath, logLines.Add);

            logLines.Should().NotBeEmpty();
            logLines.Should().Contain(line => line.Contains("Saved"));
        }
        finally
        {
            TestFiles.Delete(pptxPath);
        }
    }

    [Fact]
    public async Task WriteDocumentJsonAsync_WritesJson_WithSlideData()
    {
        var pptxPath = TestFiles.NewTempPath(".pptx");
        var jsonPath = TestFiles.NewTempPath(".pptx.json");
        try
        {
            await _sut.ExportAsync(Countries, pptxPath, _ => { });

            await _sut.WriteDocumentJsonAsync(pptxPath, jsonPath, _ => { });

            File.Exists(jsonPath).Should().BeTrue();
            File.ReadAllText(jsonPath).Should().Contain("Australia");
        }
        finally
        {
            TestFiles.Delete(pptxPath, jsonPath);
        }
    }
}