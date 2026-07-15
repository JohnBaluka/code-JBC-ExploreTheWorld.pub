using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace JBC.ExploreTheWorld.OpenXmlLibTests.Managers;

public class MsWord_OpenXml__Repo_Tests
{
    private static readonly List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS"),
        new MsOfficeCountry_Row("Brazil",    "BR", "BRA")
    };

    private readonly MsWord_OpenXml__Repo _sut = new();

    [Fact]
    public async Task ExportAsync_CreatesDocument_WithHeaderAndCountryRows()
    {
        var docxPath = TestFiles.NewTempPath(".docx");
        try
        {
            await _sut.ExportAsync(Countries, docxPath, _ => { });

            File.Exists(docxPath).Should().BeTrue();

            using var wordDoc = WordprocessingDocument.Open(docxPath, false);
            var body = wordDoc.MainDocumentPart!.Document!.Body!;

            var table = body.Elements<W.Table>().Should().ContainSingle().Subject;
            table.Elements<W.TableRow>().Should().HaveCount(Countries.Count + 1);
            body.InnerText.Should().Contain("Countries Now").And.Contain("Australia").And.Contain("BRA");
        }
        finally
        {
            TestFiles.Delete(docxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_WithFlagPng_EmbedsOneImagePartPerFlag()
    {
        var countriesWithFlags = new List<MsOfficeCountry_Row>
        {
            new("Australia", "AU", "AUS", TestFiles.TinyPng),
            new("Brazil",    "BR", "BRA", TestFiles.TinyPng),
            new("Nowhere",   "XX", "XXX") // no flag — must not add an image part
        };
        var docxPath = TestFiles.NewTempPath(".docx");
        try
        {
            await _sut.ExportAsync(countriesWithFlags, docxPath, _ => { });

            using var wordDoc = WordprocessingDocument.Open(docxPath, false);
            var mainPart = wordDoc.MainDocumentPart!;

            mainPart.ImageParts.Should().HaveCount(2);
            mainPart.Document!.Body!.Descendants<W.Drawing>().Should().HaveCount(2);

            var headerRow = mainPart.Document.Body.Elements<W.Table>().Single().Elements<W.TableRow>().First();
            headerRow.InnerText.Should().StartWith("Flag");
        }
        finally
        {
            TestFiles.Delete(docxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_ReportsProgress_ThroughLog()
    {
        var docxPath = TestFiles.NewTempPath(".docx");
        var logLines = new List<string>();
        try
        {
            await _sut.ExportAsync(Countries, docxPath, logLines.Add);

            logLines.Should().NotBeEmpty();
            logLines.Should().Contain(line => line.Contains("Saved"));
        }
        finally
        {
            TestFiles.Delete(docxPath);
        }
    }

    [Fact]
    public async Task WriteDocumentJsonAsync_WritesJson_WithParagraphsAndTables()
    {
        var docxPath = TestFiles.NewTempPath(".docx");
        var jsonPath = TestFiles.NewTempPath(".docx.json");
        try
        {
            await _sut.ExportAsync(Countries, docxPath, _ => { });

            await _sut.WriteDocumentJsonAsync(docxPath, jsonPath, _ => { });

            File.Exists(jsonPath).Should().BeTrue();
            var json = File.ReadAllText(jsonPath);
            json.Should().Contain("\"Paragraphs\"").And.Contain("\"Tables\"").And.Contain("Australia");
        }
        finally
        {
            TestFiles.Delete(docxPath, jsonPath);
        }
    }
}