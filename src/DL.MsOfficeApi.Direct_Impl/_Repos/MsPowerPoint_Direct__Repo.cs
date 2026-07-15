using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl
{
    /// <summary>
    /// "Direct" PowerPoint repository: opens the presentation in a PowerPoint host and runs
    /// the ExploreTheWorld VBA writer macro to produce the canonical JSON. Building a
    /// presentation from country data is not supported by the Direct method.
    /// </summary>
    public class MsPowerPoint_Direct__Repo : MsPowerPoint__Repo__Interface
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

                var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft PowerPoint not found. Ensure Microsoft Office is installed.");

                dynamic pptApp = Activator.CreateInstance(pptType)!;
                pptApp.Visible = -1; // msoTrue — PowerPoint is shown while the JSON is read
                try
                {
                    // The VBA macro reads ActivePresentation, which requires a window
                    // (WithWindow = msoTrue); ReadOnly = msoTrue, Untitled = msoFalse.
                    dynamic prs = pptApp.Presentations.Open(sourcePath, -1, 0, -1);
                    try
                    {
                        MsOfficeDirectVbaRunner.RunWriteMacro(
                            (object)pptApp, MsOfficeDirectVbaRunner.PowerPointMacro, outputJsonPath, log);
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
