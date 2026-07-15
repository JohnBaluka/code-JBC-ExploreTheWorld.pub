using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl
{
    /// <summary>
    /// "Direct" Word repository: opens the document in a visible Word host and runs the
    /// ExploreTheWorld VBA writer macro to produce the canonical JSON. Building a document
    /// from country data is not supported by the Direct method.
    /// </summary>
    public class MsWord_Direct__Repo : MsWord__Repo__Interface
    {
        public Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log)
            => throw new NotSupportedException(
                "The Direct method writes JSON by running the ExploreTheWorld VBA macros; it does not " +
                "build documents from country data. Use NetOffice, Dynamic, or OpenXML for export.");

        public async Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log)
        {
            await Task.Run(() =>
            {
                log($"Writing (Direct): {outputJsonPath}");

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
                        MsOfficeDirectVbaRunner.RunWriteMacro(
                            (object)wordApp, MsOfficeDirectVbaRunner.WordMacro, outputJsonPath, log);
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
