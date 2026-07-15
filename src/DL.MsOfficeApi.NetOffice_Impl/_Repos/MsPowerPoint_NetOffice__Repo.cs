using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl
{
    public class MsPowerPoint_NetOffice__Repo : MsPowerPoint__Repo__Interface, MsOfficeRunningAppExport__Repo__Interface
    {
        private const int RowsPerSlide = 15;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Launching Microsoft PowerPoint via NetOffice...");

                using var pptApp = new Application();
                pptApp.Visible = MsoTriState.msoTrue;

                pptApp.AfterNewPresentationEvent += prs =>
                    log($"  [Event] AfterNewPresentation  — {prs.Name}");
                pptApp.PresentationBeforeSaveEvent += (prs, ref cancel) =>
                    log($"  [Event] PresentationBeforeSave — {prs.Name}");

                // Create the presentation WITH a window (msoTrue). A windowless presentation
                // (msoFalse) makes PowerPoint reject SaveAs/Quit with COM errors such as
                // "Application.Quit : Invalid request. This operation cannot be performed in this
                // event handler." Word/Excel create windowed documents, which is why only
                // PowerPoint hit this.
                log("  Creating new presentation...");
                var prs2 = pptApp.Presentations.Add(MsoTriState.msoTrue);
                log($"  [Event] New presentation created: {prs2.Name}");

                FillPresentation(prs2, countries, log);

                log($"  Saving: {filePath}");
                prs2.SaveAs(filePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                log("  Closing PowerPoint...");
                pptApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        // Creates a new presentation in the caller's already-running PowerPoint application and
        // fills it, leaving it open and unsaved. Runs synchronously on the caller's (STA/UI) thread
        // — the host COM object is bound to that apartment and must not be marshalled onto a pool thread.
        public Task ExportToRunningAppAsync(object hostApplication, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            if (hostApplication is not Application pptApp)
                throw new ArgumentException(
                    "Expected a NetOffice.PowerPointApi.Application instance.", nameof(hostApplication));

            log("Creating a new presentation in the running PowerPoint application...");
            var prs = pptApp.Presentations.Add(MsoTriState.msoTrue);
            FillPresentation(prs, countries, log);
            log($"  New presentation created: {prs.Name}");
            return Task.CompletedTask;
        }

        // One country per slide (centered name, large flag, ISO codes) — mirrors the "Country Slides" page.
        private static void FillPresentation(Presentation prs, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            float slideW = prs.PageSetup.SlideWidth;
            float slideH = prs.PageSetup.SlideHeight;
            float margin = slideW * 0.06f;

            log("  Adding title slide...");
            var titleSlide = prs.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
            AddCenteredText(titleSlide, margin, slideH * 0.32f, slideW - 2 * margin, slideH * 0.22f,
                "Explore the World", 40f, bold: true);
            AddCenteredText(titleSlide, margin, slideH * 0.56f, slideW - 2 * margin, slideH * 0.12f,
                $"{countries.Count} countries", 20f);

            int slideNum = 1;
            int flagCount = 0;
            foreach (var c in countries)
            {
                slideNum++;
                var slide = prs.Slides.Add(slideNum, PpSlideLayout.ppLayoutBlank);

                AddCenteredText(slide, margin, slideH * 0.06f, slideW - 2 * margin, slideH * 0.20f,
                    c.Country, 36f, bold: true);

                var flagPath = NetOfficeFlagImage_Helper.GetFlagImageFilePath(c);
                if (flagPath != null)
                {
                    // Insert at original size, then scale by height with the aspect ratio locked and center it.
                    var picture = slide.Shapes.AddPicture(flagPath,
                        MsoTriState.msoFalse, MsoTriState.msoCTrue,
                        0f, slideH * 0.30f, -1f, -1f);
                    picture.LockAspectRatio = MsoTriState.msoTrue;
                    picture.Height = slideH * 0.42f;
                    picture.Left   = (slideW - picture.Width) / 2f;
                    flagCount++;
                }

                var codes = string.Join("      ",
                    new[] { c.Iso2, c.Iso3 }.Where(s => !string.IsNullOrEmpty(s)));
                AddCenteredText(slide, margin, slideH * 0.80f, slideW - 2 * margin, slideH * 0.12f, codes, 20f);

                if (slideNum % 25 == 0)
                    log($"  Added {slideNum} slides...");
            }
            log($"  Built {slideNum} slides; embedded {flagCount} flag images.");
        }

        private static void AddCenteredText(
            Slide slide, float left, float top, float width, float height, string text, float fontSize, bool bold = false)
        {
            var box = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            var range = box.TextFrame.TextRange;
            range.Text      = text;
            range.Font.Size = fontSize;
            range.Font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            range.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            box.TextFrame.VerticalAnchor    = MsoVerticalAnchor.msoAnchorMiddle;
            box.TextFrame.WordWrap          = MsoTriState.msoTrue;
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (NetOffice): {outputJsonPath}");

                // PowerPoint is shown while the JSON is read: Application.Visible = msoTrue
                // and the presentation is opened WithWindow=msoTrue so it surfaces on screen.
                using var pptApp = new Application();
                pptApp.Visible = MsoTriState.msoTrue;
                var prs = pptApp.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);
                try
                {
                    WriteJsonFromOpenPresentation(prs, outputJsonPath, log);
                }
                finally
                {
                    prs.Close();
                    pptApp.Quit();
                }
            });
        }

        public void WriteJsonFromOpenPresentation(Presentation prs, string outputJsonPath, Action<string> log)
            => JsonWriters.MsPowerPointJsonWriter.WritePresentationToJsonFile(prs, outputJsonPath, null, log);
    }
}