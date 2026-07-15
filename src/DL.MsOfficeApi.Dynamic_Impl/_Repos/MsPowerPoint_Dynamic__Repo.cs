using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl
{
    public class MsPowerPoint_Dynamic__Repo : MsPowerPoint__Repo__Interface
    {
        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Connecting to Microsoft PowerPoint via dynamic COM automation...");

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft PowerPoint not found. Ensure Microsoft Office is installed.");

                log("  Creating PowerPoint application instance...");
                dynamic pptApp = Activator.CreateInstance(pptType)!;
                pptApp.Visible = 1; // msoTrue
                log($"  [Event] Connected — PowerPoint {pptApp.Version}");

                // Create the presentation WITH a window (msoTrue = -1). A windowless presentation
                // makes PowerPoint reject Close/Quit with COM errors such as "Invalid request. This
                // operation cannot be performed in this event handler." Word/Excel create windowed
                // documents, which is why only PowerPoint hit this.
                log("  Creating new presentation...");
                dynamic prs = pptApp.Presentations.Add(-1); // msoTrue = with window
                log($"  [Event] AfterNewPresentation — {prs.Name}");

                float slideW = Convert.ToSingle(prs.PageSetup.SlideWidth);
                float slideH = Convert.ToSingle(prs.PageSetup.SlideHeight);
                float margin = slideW * 0.06f;

                // One country per slide (centered name, large flag, ISO codes) — like the "Country Slides" page.
                log("  Adding title slide...");
                dynamic titleSlide = prs.Slides.Add(1, 12); // ppLayoutBlank = 12
                AddCenteredText(titleSlide, margin, slideH * 0.32f, slideW - 2 * margin, slideH * 0.22f,
                    "Explore the World", 40f, bold: true);
                AddCenteredText(titleSlide, margin, slideH * 0.56f, slideW - 2 * margin, slideH * 0.12f,
                    $"{countries.Count} countries", 20f);

                int slideNum = 1;
                int flagCount = 0;
                foreach (var c in countries)
                {
                    slideNum++;
                    dynamic slide = prs.Slides.Add(slideNum, 12); // ppLayoutBlank = 12

                    AddCenteredText(slide, margin, slideH * 0.06f, slideW - 2 * margin, slideH * 0.20f,
                        c.Country, 36f, bold: true);

                    var flagPath = DynamicFlagImage_Helper.GetFlagImageFilePath(c);
                    if (flagPath != null)
                    {
                        // Insert at original size, then scale by height with aspect ratio locked and center it.
                        dynamic picture = slide.Shapes.AddPicture(flagPath,
                            0, 1, 0f, slideH * 0.30f, -1f, -1f); // msoFalse, msoCTrue
                        picture.LockAspectRatio = -1; // msoTrue
                        picture.Height = slideH * 0.42f;
                        picture.Left   = (slideW - Convert.ToSingle(picture.Width)) / 2f;
                        flagCount++;
                    }

                    var codes = string.Join("      ",
                        new[] { c.Iso2, c.Iso3 }.Where(s => !string.IsNullOrEmpty(s)));
                    AddCenteredText(slide, margin, slideH * 0.80f, slideW - 2 * margin, slideH * 0.12f, codes, 20f);

                    if (slideNum % 25 == 0)
                        log($"  Added {slideNum} slides...");
                }
                log($"  Built {slideNum} slides; embedded {flagCount} flag images.");

                log($"  [Event] PresentationBeforeSave — {prs.Name}");
                log($"  Saving: {filePath}");
                prs.SaveAs(filePath, 24); // ppSaveAsOpenXMLPresentation = 24
                log("  [Event] Presentation saved.");
                prs.Close();
                pptApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        // Adds a horizontally- and vertically-centered text box. Enum literals are used because
        // dynamic COM has no typed enums: ppAlignCenter = 2, msoAnchorMiddle = 3,
        // msoTextOrientationHorizontal = 1, msoTrue = -1, msoFalse = 0.
        private static void AddCenteredText(
            dynamic slide, float left, float top, float width, float height, string text, float fontSize, bool bold = false)
        {
            dynamic box = slide.Shapes.AddTextbox(1, left, top, width, height); // msoTextOrientationHorizontal
            dynamic range = box.TextFrame.TextRange;
            range.Text      = text;
            range.Font.Size = fontSize;
            range.Font.Bold = bold ? -1 : 0; // msoTrue / msoFalse
            range.ParagraphFormat.Alignment = 2; // ppAlignCenter
            box.TextFrame.VerticalAnchor    = 3; // msoAnchorMiddle
            box.TextFrame.WordWrap          = -1; // msoTrue
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (Dynamic): {outputJsonPath}");

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft PowerPoint not found. Ensure Microsoft Office is installed.");

                dynamic pptApp = Activator.CreateInstance(pptType)!;
                pptApp.Visible = -1; // msoTrue — PowerPoint is shown while the JSON is read
                try
                {
                    // ReadOnly = msoTrue, Untitled = msoTrue, WithWindow = msoTrue
                    dynamic prs = pptApp.Presentations.Open(sourcePath, -1, -1, -1);
                    try
                    {
                        JsonWriters.MsPowerPointJsonWriter.WritePresentationComToJsonFile(prs, outputJsonPath, null, log);
                    }
                    finally
                    {
                        prs.Close();
                    }
                }
                finally
                {
                    pptApp.Quit();
                }
            });
        }
    }
}
