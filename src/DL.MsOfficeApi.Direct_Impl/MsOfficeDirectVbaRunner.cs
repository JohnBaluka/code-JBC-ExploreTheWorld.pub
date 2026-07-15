using System;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl
{
    /// <summary>
    /// Runs the ExploreTheWorld VBA "Save as JSON" writer macros inside a live Office host
    /// through late-bound dynamic COM (<c>Application.Run</c>). This is the "Direct" option:
    /// instead of walking the COM object model in C# (the "Dynamic" method), it delegates to
    /// the VBA <c>MSO_Ms{Host}.WriteActive…ToJsonFile</c> procedure that must already be
    /// loaded in the host (for example ExploreTheWorld.dotm in Word's STARTUP folder).
    /// <para>
    /// No NetOffice dependency: callers pass the raw COM application object — NetOffice's
    /// <c>UnderlyingObject</c>, an Interop PIA application, or a freshly created late-bound
    /// instance — and the runner dispatches through <c>dynamic</c>.
    /// </para>
    /// </summary>
    public static class MsOfficeDirectVbaRunner
    {
        /// <summary>VBA writer macro procedure names, addressed as Module.Procedure.</summary>
        public const string WordMacro       = "MSO_MsWord.WriteActiveDocumentToJsonFile";
        public const string ExcelMacro      = "MSO_MsExcel.WriteActiveWorkbookToJsonFile";
        public const string PowerPointMacro = "MSO_MsPowerPoint.WriteActivePresentationToJsonFile";

        /// <summary>
        /// Runs a writer macro on an already-open Office application COM object, passing the
        /// output JSON path as the single argument. The macro reads the host's active
        /// document/workbook/presentation and writes the canonical JSON to
        /// <paramref name="outputJsonPath"/>.
        /// </summary>
        /// <param name="app">The Office application COM object (late-bound or a NetOffice UnderlyingObject).</param>
        /// <param name="macro">The macro procedure, addressed as Module.Procedure.</param>
        /// <param name="outputJsonPath">The output .json path passed to the macro.</param>
        /// <param name="log">Progress/diagnostic sink.</param>
        public static void RunWriteMacro(object app, string macro, string outputJsonPath, Action<string> log)
        {
            if (app is null) throw new ArgumentNullException(nameof(app));

            try
            {
                ((dynamic)app).Run(macro, outputJsonPath);
                log($"Direct write complete: {outputJsonPath}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "The Direct method requires the ExploreTheWorld VBA macros to be loaded in the Office " +
                    "host (see the VBA folder in the repository). The macro call failed: " + ex.Message, ex);
            }
        }
    }
}
