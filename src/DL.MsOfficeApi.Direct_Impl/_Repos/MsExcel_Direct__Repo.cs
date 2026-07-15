using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl
{
    /// <summary>
    /// "Direct" Excel repository: opens the workbook in a visible Excel host and runs the
    /// ExploreTheWorld VBA writer macro to produce the canonical JSON. Building a workbook
    /// from country data is not supported by the Direct method.
    /// </summary>
    public class MsExcel_Direct__Repo : MsExcel__Repo__Interface
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

                var excelType = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Excel not found. Ensure Microsoft Office is installed.");

                dynamic excelApp = Activator.CreateInstance(excelType)!;
                excelApp.Visible       = true;
                excelApp.DisplayAlerts = false;
                try
                {
                    dynamic wb = excelApp.Workbooks.Open(sourcePath);
                    try
                    {
                        MsOfficeDirectVbaRunner.RunWriteMacro(
                            (object)excelApp, MsOfficeDirectVbaRunner.ExcelMacro, outputJsonPath, log);
                    }
                    finally
                    {
                        wb.Close(false);
                    }
                }
                finally
                {
                    excelApp.Quit();
                }
            });
        }
    }
}
