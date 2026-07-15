using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using NetOffice.WordApi;
using Task = System.Threading.Tasks.Task;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl
{
    public class MsWord_NetOffice__Repo : MsWord__Repo__Interface, MsOfficeRunningAppExport__Repo__Interface
    {
        private const float FlagHeightPoints = 18f;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Launching Microsoft Word via NetOffice...");

                using var wordApp = new Application { Visible = true };

                wordApp.DocumentBeforeCloseEvent += (Document doc, ref bool cancel) =>
                    log($"  [Event] DocumentBeforeClose  — {doc.Name}");
                wordApp.DocumentBeforeSaveEvent += (Document doc, ref bool saveAsUi, ref bool cancel) =>
                    log($"  [Event] DocumentBeforeSave   — {doc.Name}");
                wordApp.DocumentOpenEvent += doc =>
                    log($"  [Event] DocumentOpen         — {doc.Name}");

                log("  Creating new document...");
                var doc2 = wordApp.Documents.Add();
                log($"  [Event] New document created: {doc2.Name}");

                FillDocument(doc2, countries, log);

                log($"  Saving: {filePath}");
                doc2.SaveAs2(filePath);
                log("  Closing Word...");
                wordApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        // Creates a new document in the caller's already-running Word application and fills it,
        // leaving it open and unsaved. Runs synchronously on the caller's (STA/UI) thread — the
        // host COM object is bound to that apartment and must not be marshalled onto a pool thread.
        public Task ExportToRunningAppAsync(object hostApplication, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            if (hostApplication is not Application wordApp)
                throw new ArgumentException(
                    "Expected a NetOffice.WordApi.Application instance.", nameof(hostApplication));

            log("Creating a new document in the running Word application...");
            var doc = wordApp.Documents.Add();
            FillDocument(doc, countries, log);
            log($"  New document created: {doc.Name}");
            return Task.CompletedTask;
        }

        private static void FillDocument(Document doc, IList<MsOfficeCountry_Row> countries, Action<string> log)
        {
            log("  Adding title...");
            doc.Range().Text = "Countries Now\n\n";

            log($"  Adding table (4 columns, {countries.Count} rows)...");
            var range = doc.Content;
            range.Collapse(NetOffice.WordApi.Enums.WdCollapseDirection.wdCollapseEnd);
            var table = doc.Tables.Add(range, countries.Count + 1, 4);

            table.Cell(1, 1).Range.Text = "Flag";
            table.Cell(1, 2).Range.Text = "Country";
            table.Cell(1, 3).Range.Text = "ISO2";
            table.Cell(1, 4).Range.Text = "ISO3";

            int r = 0;
            int flagCount = 0;
            foreach (var c in countries)
            {
                r++;
                var flagPath = NetOfficeFlagImage_Helper.GetFlagImageFilePath(c);
                if (flagPath != null)
                {
                    var picture = doc.InlineShapes.AddPicture(
                        flagPath, false, true, table.Cell(r + 1, 1).Range);
                    picture.LockAspectRatio = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue; // scale width with height
                    picture.Height = FlagHeightPoints;
                    flagCount++;
                }
                table.Cell(r + 1, 2).Range.Text = c.Country;
                table.Cell(r + 1, 3).Range.Text = c.Iso2;
                table.Cell(r + 1, 4).Range.Text = c.Iso3;
                if (r % 50 == 0)
                    log($"  Written {r}/{countries.Count} rows...");
            }
            log($"  Embedded {flagCount} flag images.");
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (NetOffice): {outputJsonPath}");

                using var wordApp = new Application { Visible = true };
                var doc = wordApp.Documents.Open(sourcePath);
                try
                {
                    WriteJsonFromOpenDocument(doc, outputJsonPath, log);
                }
                finally
                {
                    doc.Close(NetOffice.WordApi.Enums.WdSaveOptions.wdDoNotSaveChanges);
                    wordApp.Quit();
                }
            });
        }

        public void WriteJsonFromOpenDocument(Document doc, string outputJsonPath, Action<string> log)
            => JsonWriters.MsWordJsonWriter.WriteDocumentToJsonFile(doc, outputJsonPath, null, log);
    }
}