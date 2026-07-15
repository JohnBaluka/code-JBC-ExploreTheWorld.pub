using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace JBC.ExploreTheWorld.OpenXmlLibTests.Managers;

public class MsExcel_OpenXml__Repo_Tests
{
    private static readonly List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS"),
        new MsOfficeCountry_Row("Brazil",    "BR", "BRA")
    };

    private readonly MsExcel_OpenXml__Repo _sut = new();

    [Fact]
    public async Task ExportAsync_CreatesWorkbook_WithCountriesSheet()
    {
        var xlsxPath = TestFiles.NewTempPath(".xlsx");
        try
        {
            await _sut.ExportAsync(Countries, xlsxPath, _ => { });

            File.Exists(xlsxPath).Should().BeTrue();

            using var spreadsheet = SpreadsheetDocument.Open(xlsxPath, false);
            var workbookPart = spreadsheet.WorkbookPart!;

            var sheet = workbookPart.Workbook!.Descendants<X.Sheet>().Should().ContainSingle().Subject;
            sheet.Name!.Value.Should().Be("Countries");

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
            var rows = worksheetPart.Worksheet!.Descendants<X.Row>();
            rows.Should().HaveCount(Countries.Count + 1, "one header row plus one row per country");
            worksheetPart.Worksheet.InnerText.Should().Contain("Australia").And.Contain("BRA");
        }
        finally
        {
            TestFiles.Delete(xlsxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_WithFlagPng_EmbedsAnchoredImagesInDrawingsPart()
    {
        var countriesWithFlags = new List<MsOfficeCountry_Row>
        {
            new("Australia", "AU", "AUS", TestFiles.TinyPng),
            new("Brazil",    "BR", "BRA", TestFiles.TinyPng),
            new("Nowhere",   "XX", "XXX") // no flag — must not add an image part
        };
        var xlsxPath = TestFiles.NewTempPath(".xlsx");
        try
        {
            await _sut.ExportAsync(countriesWithFlags, xlsxPath, _ => { });

            using var spreadsheet = SpreadsheetDocument.Open(xlsxPath, false);
            var workbookPart  = spreadsheet.WorkbookPart!;
            var sheet         = workbookPart.Workbook!.Descendants<X.Sheet>().Single();
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);

            worksheetPart.DrawingsPart.Should().NotBeNull();
            worksheetPart.DrawingsPart!.ImageParts.Should().HaveCount(2);
            worksheetPart.DrawingsPart.WorksheetDrawing!
                .Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>()
                .Should().HaveCount(2);

            var headerRow = worksheetPart.Worksheet!.Descendants<X.Row>().First();
            headerRow.InnerText.Should().StartWith("Flag");
        }
        finally
        {
            TestFiles.Delete(xlsxPath);
        }
    }

    [Fact]
    public async Task ExportAsync_ReportsProgress_ThroughLog()
    {
        var xlsxPath = TestFiles.NewTempPath(".xlsx");
        var logLines = new List<string>();
        try
        {
            await _sut.ExportAsync(Countries, xlsxPath, logLines.Add);

            logLines.Should().NotBeEmpty();
            logLines.Should().Contain(line => line.Contains("Saved"));
        }
        finally
        {
            TestFiles.Delete(xlsxPath);
        }
    }

    [Fact]
    public async Task WriteDocumentJsonAsync_WritesJson_WithCountryData()
    {
        var xlsxPath = TestFiles.NewTempPath(".xlsx");
        var jsonPath = TestFiles.NewTempPath(".xlsx.json");
        try
        {
            await _sut.ExportAsync(Countries, xlsxPath, _ => { });

            await _sut.WriteDocumentJsonAsync(xlsxPath, jsonPath, _ => { });

            File.Exists(jsonPath).Should().BeTrue();
            File.ReadAllText(jsonPath).Should().Contain("Australia");
        }
        finally
        {
            TestFiles.Delete(xlsxPath, jsonPath);
        }
    }
}