using System.Collections.Generic;
using System.Linq;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;
using WD = JBC.ExploreTheWorld.DL.MsOfficeApi.MsWord;
using XL = JBC.ExploreTheWorld.DL.MsOfficeApi.MsExcel;

namespace JBC.ExploreTheWorld.OpenXmlLibTests.JsonWriters;

// The OpenXML "Save as JSON" writers must emit the canonical entity schema so the
// output deserializes back into the JBC.ExploreTheWorld.DL entities.
public class MsOfficeJsonWriters_Tests
{
    private static readonly List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS"),
        new MsOfficeCountry_Row("Brazil",    "BR", "BRA")
    };

    [Fact]
    public async Task PowerPointJson_DeserializesIntoCanonicalEntities()
    {
        var repo = new MsPowerPoint_OpenXml__Repo();
        var pptxPath = TestFiles.NewTempPath(".pptx");
        var jsonPath = TestFiles.NewTempPath(".pptx.json");
        try
        {
            await repo.ExportAsync(Countries, pptxPath, _ => { });
            await repo.WriteDocumentJsonAsync(pptxPath, jsonPath, _ => { });

            var presentation = MsOfficeJsonSerializer.ReadFromFile<PP.Presentation>(jsonPath);

            presentation.Should().NotBeNull();
            presentation!.Name.Should().Be(Path.GetFileName(pptxPath));
            presentation.Slides.Should().HaveCount(Countries.Count + 1, "a title slide plus one slide per country");
            presentation.Slides![0].SlideIndex.Should().Be(1);
            presentation.Slides.SelectMany(s => s.Shapes ?? new List<PP.Shape>())
                .Any(shape => shape.TextFrame?.TextRange?.Text?.Contains("Australia") == true)
                .Should().BeTrue();

            // VBA-only properties are marked with the canonical Undefined values.
            MsOfficeUndefined.IsUndefined(presentation.GridDistance).Should().BeTrue();
            MsOfficeUndefined.IsUndefined(presentation.DefaultLanguageID).Should().BeTrue();

            // The gold (COM) writer reports ppMediaTaskStatusNone for any file on disk.
            presentation.CreateVideoStatus.Should().Be(0);
        }
        finally
        {
            TestFiles.Delete(pptxPath, jsonPath);
        }
    }

    [Fact]
    public async Task ExcelJson_DeserializesIntoCanonicalEntities()
    {
        var repo = new MsExcel_OpenXml__Repo();
        var xlsxPath = TestFiles.NewTempPath(".xlsx");
        var jsonPath = TestFiles.NewTempPath(".xlsx.json");
        try
        {
            await repo.ExportAsync(Countries, xlsxPath, _ => { });
            await repo.WriteDocumentJsonAsync(xlsxPath, jsonPath, _ => { });

            var workbook = MsOfficeJsonSerializer.ReadFromFile<XL.Workbook>(jsonPath);

            workbook.Should().NotBeNull();
            workbook!.FileFormat.Should().Be(51, "an .xlsx file is xlOpenXMLWorkbook");
            workbook.Sheets.Should().HaveCount(1);

            var sheet = workbook.Sheets![0];
            sheet.Type.Should().Be("Worksheet");
            sheet.UsedRange!.RowCount.Should().Be(3, "a header row plus two data rows");
            sheet.Rows![0].Cells![0].Value.Should().Be("Flag", "column A holds the flag images");
            sheet.Rows[0].Cells![1].Value.Should().Be("Country");
            sheet.Rows[1].Cells![1].Value.Should().Be("Australia");
        }
        finally
        {
            TestFiles.Delete(xlsxPath, jsonPath);
        }
    }

    [Fact]
    public async Task WordJson_DeserializesIntoCanonicalEntities()
    {
        var repo = new MsWord_OpenXml__Repo();
        var docxPath = TestFiles.NewTempPath(".docx");
        var jsonPath = TestFiles.NewTempPath(".docx.json");
        try
        {
            await repo.ExportAsync(Countries, docxPath, _ => { });
            await repo.WriteDocumentJsonAsync(docxPath, jsonPath, _ => { });

            var document = MsOfficeJsonSerializer.ReadFromFile<WD.Document>(jsonPath);

            document.Should().NotBeNull();
            document!.Name.Should().Be(Path.GetFileName(docxPath));
            document.Paragraphs.Should().NotBeEmpty();
            document.Paragraphs!.Any(p => p.Text?.Contains("Countries Now") == true).Should().BeTrue();
            document.Tables.Should().HaveCount(1);
            document.Tables![0].Rows![0].Cells!.Select(c => c.Text)
                .Should().Contain(text => text != null && text.Contains("Country"));
        }
        finally
        {
            TestFiles.Delete(docxPath, jsonPath);
        }
    }
}
