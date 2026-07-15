using System.Diagnostics;
using JBC.ExploreTheWorld.DL.MsOfficeApi_Impl;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.CL;

namespace JBC.ExploreTheWorld.AL.SaveAsJson.ConsoleApp
{
    /// <summary>
    /// Console version of the Watcher forms' "Save As JSON": writes the canonical
    /// JBC.ExploreTheWorld.DL.Ms{Word|Excel|PowerPoint} JSON representation of an Office
    /// document with the selected write method (NetOffice / OpenXml / Dynamic / Direct).
    /// </summary>
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            var options = SaveAsJsonOptions.Parse(args);

            if (options.ShowHelp)
            {
                Console.WriteLine(SaveAsJsonOptions.Usage);
                return 0;
            }

            if (options.Error != null)
            {
                Console.Error.WriteLine(options.Error);
                Console.Error.WriteLine();
                Console.Error.WriteLine(SaveAsJsonOptions.Usage);
                return 1;
            }

            Log($"Source : {options.SourcePath}");
            Log($"Output : {options.OutputPath}");
            Log($"Save As JSON started ({options.Type} — {options.Method})");

            var stopwatch = Stopwatch.StartNew();
            try
            {
                // Every method — NetOffice, OpenXml, Dynamic, and Direct — is routed through the
                // export manager and the host factory. The "Direct" repos run the ExploreTheWorld
                // VBA writer macros in the Office host (DL.MsOfficeApi.Direct_Impl).
                var exportManager = new MsOfficeExportManager__Service(new MsOfficeExportRepoFactory());
                await exportManager.WriteDocumentJsonAsync(
                    options.Type, options.Method, options.SourcePath, options.OutputPath, Log);

                stopwatch.Stop();
                Log($"✔ Save As JSON complete ({options.Method}) in {Duration_Helper.Format(stopwatch.Elapsed)} → {options.OutputPath}");
                return 0;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Log($"✘ Save As JSON failed after {Duration_Helper.Format(stopwatch.Elapsed)}: {ex.Message}");
                return 1;
            }
        }

        private static void Log(string message)
            => Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
    }
}
