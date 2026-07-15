using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl
{
    public class MsWord_Dynamic__Repo : MsWord__Repo__Interface
    {
        private const float FlagHeightPoints = 18f;

        public async Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log("Connecting to Microsoft Word via dynamic COM automation...");

                var wordType = Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Word not found. Ensure Microsoft Office is installed.");

                log("  Creating Word application instance...");
                dynamic wordApp = Activator.CreateInstance(wordType)!;
                wordApp.Visible = true; // keep the app visible while exporting
                log($"  [Event] Connected — Word {wordApp.Version}");

                log("  Creating new document...");
                dynamic doc = wordApp.Documents.Add();
                log($"  [Event] DocumentOpen — {doc.Name}");

                log("  Adding title...");
                doc.Content.Text = "Countries Now\n\n";

                log($"  Adding table (4 columns, {countries.Count} rows)...");
                dynamic range = doc.Content;
                range.Collapse(0); // wdCollapseEnd
                dynamic table = doc.Tables.Add(range, countries.Count + 1, 4);

                table.Cell(1, 1).Range.Text = "Flag";
                table.Cell(1, 2).Range.Text = "Country";
                table.Cell(1, 3).Range.Text = "ISO2";
                table.Cell(1, 4).Range.Text = "ISO3";

                int r = 0;
                int flagCount = 0;
                foreach (var c in countries)
                {
                    r++;
                    var flagPath = DynamicFlagImage_Helper.GetFlagImageFilePath(c);
                    if (flagPath != null)
                    {
                        dynamic picture = doc.InlineShapes.AddPicture(
                            flagPath, false, true, table.Cell(r + 1, 1).Range);
                        picture.LockAspectRatio = -1; // msoTrue — scale width with height
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

                log($"  [Event] DocumentBeforeSave — {doc.Name}");
                log($"  Saving: {filePath}");
                doc.SaveAs2(filePath);
                log("  [Event] Document saved.");
                doc.Close();
                wordApp.Quit();
                log($"  Saved: {filePath}");
            });
        }

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (Dynamic): {outputJsonPath}");

                var wordType = Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Word not found. Ensure Microsoft Office is installed.");

                dynamic wordApp = Activator.CreateInstance(wordType)!;
                wordApp.Visible = true;
                try
                {
                    dynamic doc = wordApp.Documents.Open(sourcePath);
                    try
                    {
                        JsonWriters.MsWordJsonWriter.WriteDocumentComToJsonFile(doc, outputJsonPath, null, log);
                    }
                    finally
                    {
                        doc.Close(0); // wdDoNotSaveChanges = 0
                    }
                }
                finally
                {
                    wordApp.Quit();
                }
            });
        }
    }
}
