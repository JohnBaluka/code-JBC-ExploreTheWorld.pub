using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl
{
    public class MsPowerPoint_OpenXml__Repo : MsPowerPoint__Repo__Interface
    {
        // One country per slide, centered — mirrors the Blazor "Country Slides" look
        // (big flag, country name, ISO codes). Positions are in EMU on the 9,144,000 × 5,143,500 slide.
        private const long SlideWidthEmu  = 9144000;
        private const long MarginEmu      = 457200;
        private const long NameTopEmu     = 350000;
        private const long NameHeightEmu  = 900000;
        private const long FlagTopEmu     = 1400000;
        private const long FlagHeightEmu  = 2200000;
        private const long IsoTopEmu      = 4150000;
        private const long IsoHeightEmu   = 600000;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Creating PowerPoint presentation (DocumentFormat.OpenXml)...");

                using var prs = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
                var prsPart = prs.AddPresentationPart();
                // ECMA-376 child order: sldMasterIdLst, sldIdLst, sldSz, notesSz — PowerPoint
                // silently drops all slides when the id lists come after the size elements.
                var presentation = new P.Presentation(
                    new P.SlideMasterIdList(),
                    new P.SlideIdList(),
                    new P.SlideSize { Cx = 9144000, Cy = 5143500 },
                    new P.NotesSize { Cx = 6858000, Cy = 9144000 });
                prsPart.Presentation = presentation;
                log("  PresentationPart created.");

                log("  Building slide master and layout...");
                InitializeMaster(prsPart);

                log("  Adding title slide...");
                AddTitleSlide(prsPart, "Explore the World", $"{countries.Count} countries");

                int slideNum = 1;
                int flagCount = 0;
                foreach (var country in countries)
                {
                    slideNum++;
                    AddCountrySlide(prsPart, country, ref flagCount);
                    if (slideNum % 25 == 0)
                        log($"  Added {slideNum} slides...");
                }

                log($"  Built {slideNum} slides; embedded {flagCount} flag images.");
                log("  Saving presentation...");
                prsPart.Presentation.Save();
                log($"  Saved: {filePath}");
            });
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
                JsonWriters.MsPowerPointJsonWriter.WritePresentationFileToJson(sourcePath, outputJsonPath, null, log));
        }
        // ── Slide builder helpers ─────────────────────────────────────────────

        private static void InitializeMaster(PresentationPart prsPart)
        {
            var smPart     = prsPart.AddNewPart<SlideMasterPart>();
            var layoutPart = smPart.AddNewPart<SlideLayoutPart>();

            // PowerPoint reports the file as corrupted unless the layout references its
            // master back and the master has a theme part (also referenced by the
            // presentation part), like every PowerPoint-saved package.
            layoutPart.AddPart(smPart);

            var themePart = smPart.AddNewPart<ThemePart>();
            themePart.Theme = OpenXmlMinimalTheme.Create();
            themePart.Theme.Save();
            prsPart.AddPart(themePart);

            layoutPart.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(new D.TransformGroup()))),
                new P.ColorMapOverride(new D.MasterColorMapping()));
            layoutPart.SlideLayout.Save();

            smPart.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(new D.TransformGroup()))),
                new P.ColorMap
                {
                    Background1       = D.ColorSchemeIndexValues.Light1,
                    Text1             = D.ColorSchemeIndexValues.Dark1,
                    Background2       = D.ColorSchemeIndexValues.Light2,
                    Text2             = D.ColorSchemeIndexValues.Dark2,
                    Accent1           = D.ColorSchemeIndexValues.Accent1,
                    Accent2           = D.ColorSchemeIndexValues.Accent2,
                    Accent3           = D.ColorSchemeIndexValues.Accent3,
                    Accent4           = D.ColorSchemeIndexValues.Accent4,
                    Accent5           = D.ColorSchemeIndexValues.Accent5,
                    Accent6           = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink         = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new P.SlideLayoutIdList(
                    // ST_SlideLayoutId requires ids >= 0x80000000.
                    new P.SlideLayoutId { Id = 2147483649U, RelationshipId = smPart.CreateRelationshipToPart(layoutPart) }));
            smPart.SlideMaster.Save();

            string masterRelId = prsPart.CreateRelationshipToPart(smPart);
            prsPart.Presentation.GetFirstChild<P.SlideMasterIdList>()!
                   .AppendChild(new P.SlideMasterId { Id = 2147483648U, RelationshipId = masterRelId });
        }

        private static SlidePart AddSlidePart(PresentationPart prsPart, P.ShapeTree shapeTree)
        {
            var slidePart       = prsPart.AddNewPart<SlidePart>();
            var slideLayoutPart = prsPart.SlideMasterParts.First().SlideLayoutParts.First();
            slidePart.AddPart(slideLayoutPart);

            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(shapeTree),
                new P.ColorMapOverride(new D.MasterColorMapping()));
            slidePart.Slide.Save();

            string relId    = prsPart.CreateRelationshipToPart(slidePart);
            var slideIdList = prsPart.Presentation.GetFirstChild<P.SlideIdList>()!;
            uint nextId     = 256U;
            if (slideIdList.ChildElements.Count > 0)
                nextId = slideIdList.OfType<P.SlideId>().Max(s => (uint)s.Id!) + 1;
            slideIdList.AppendChild(new P.SlideId { Id = nextId, RelationshipId = relId });
            return slidePart;
        }

        // Title slide: centered heading + country count (no export timestamp).
        private static void AddTitleSlide(PresentationPart prsPart, string title, string subtitle)
        {
            AddSlidePart(prsPart, new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new D.TransformGroup()),
                BuildTextShape(2, "Title",    MarginEmu, 1600000, SlideWidthEmu - 2 * MarginEmu, 1200000, title,    4400, bold: true, center: true),
                BuildTextShape(3, "Subtitle", MarginEmu, 2900000, SlideWidthEmu - 2 * MarginEmu,  700000, subtitle, 2000, center: true)));
        }

        // One country per slide: centered name, large flag, ISO codes — like the "Country Slides" page.
        private static void AddCountrySlide(
            PresentationPart prsPart, MsOfficeCountry_Row country, ref int flagCount)
        {
            var shapeTree = new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new D.TransformGroup()),
                BuildTextShape(2, "Name", MarginEmu, NameTopEmu, SlideWidthEmu - 2 * MarginEmu, NameHeightEmu,
                    country.Country, 3600, bold: true, center: true));

            var slidePart = AddSlidePart(prsPart, shapeTree);
            var tree      = slidePart.Slide.CommonSlideData!.ShapeTree!;
            uint shapeId  = 3;

            if (country.FlagPng is { Length: > 0 })
            {
                var imagePart = slidePart.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream(country.FlagPng))
                    imagePart.FeedData(stream);

                var widthEmu = PngImage_Helper.GetWidthForHeight(country.FlagPng, FlagHeightEmu);
                var flagLeft = (SlideWidthEmu - widthEmu) / 2; // centered horizontally
                tree.AppendChild(OpenXmlFlagImage_Helper.CreatePresentationPicture(
                    slidePart.GetIdOfPart(imagePart), shapeId++, $"Flag {country.Iso2}",
                    flagLeft, FlagTopEmu, widthEmu, FlagHeightEmu));
                flagCount++;
            }

            var codes = string.Join("      ",
                new[] { country.Iso2, country.Iso3 }.Where(s => !string.IsNullOrEmpty(s)));
            tree.AppendChild(BuildTextShape(shapeId++, "Codes",
                MarginEmu, IsoTopEmu, SlideWidthEmu - 2 * MarginEmu, IsoHeightEmu, codes, 2000, center: true));

            slidePart.Slide.Save();
        }

        private static P.Shape BuildTextShape(
            uint shapeId, string name, long x, long y, long cx, long cy, string text, int fontSize,
            bool bold = false, bool center = false)
        {
            var shape = new P.Shape();
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape()));
            shape.ShapeProperties = new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset  { X = x, Y = y },
                    new D.Extents { Cx = cx, Cy = cy }));

            var bodyProps = new D.BodyProperties();
            if (center)
                bodyProps.Anchor = D.TextAnchoringTypeValues.Center; // vertical centering
            var txBody = new P.TextBody(bodyProps, new D.ListStyle());
            foreach (var line in text.Split('\n'))
            {
                var paragraph = new D.Paragraph();
                if (center)
                    paragraph.AppendChild(new D.ParagraphProperties { Alignment = D.TextAlignmentTypeValues.Center });
                paragraph.AppendChild(new D.Run(
                    new D.RunProperties { Language = "en-US", FontSize = fontSize, Bold = bold }, new D.Text(line)));
                txBody.AppendChild(paragraph);
            }
            shape.TextBody = txBody;
            return shape;
        }

    }
}
