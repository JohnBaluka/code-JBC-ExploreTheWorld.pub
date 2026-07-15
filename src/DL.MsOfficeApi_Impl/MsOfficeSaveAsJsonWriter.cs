using System;
using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.Direct_Impl;
using Word = NetOffice.WordApi;
using Excel = NetOffice.ExcelApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    /// <summary>
    /// Host-side implementation of <see cref="MsOfficeSaveAsJsonWriter__Interface"/>. Dispatches the
    /// user-selected write method to the platform-specific DL writers (the single place, alongside
    /// <see cref="MsOfficeExportRepoFactory"/>, permitted to reference the DL repo `_Impl` projects).
    /// The watcher forms in the AL WinForms UI libraries reach it through
    /// <c>MsOfficeSaveAsJsonWriterProvider</c> so those libraries reference no `_Impl` project.
    /// </summary>
    public sealed class MsOfficeSaveAsJsonWriter : MsOfficeSaveAsJsonWriter__Interface
    {
        public void WriteWordJson(string method, object? document, object? application, string outputFilePath, Action<string> log)
        {
            switch (method)
            {
                case "NetOffice": MsWord_JsonWriter.WriteNetOffice(ReqDoc(document), outputFilePath, log); break;
                case "Interop":   MsWord_JsonWriter.WriteInterop(ReqDoc(document), outputFilePath, log); break;
                case "Dynamic":   MsWord_JsonWriter.WriteDynamic(ReqDoc(document), outputFilePath, log); break;
                case "OpenXml":   MsWord_JsonWriter.WriteOpenXmlFromRunningApp(ReqApp(application), outputFilePath, log); break;
                case "Direct":    MsOfficeDirectVbaRunner.RunWriteMacro(
                                      ReqApp(application).UnderlyingObject!, MsOfficeDirectVbaRunner.WordMacro, outputFilePath, log); break;
                default:          log($"Unknown method: {method}"); break;
            }

            static Word.Document ReqDoc(object? o) =>
                o as Word.Document ?? throw new InvalidOperationException("No active document.");
            static Word.Application ReqApp(object? o) =>
                o as Word.Application ?? throw new InvalidOperationException("No Word connection.");
        }

        public void WriteExcelJson(string method, object? workbook, object? application, string outputFilePath, Action<string> log)
        {
            switch (method)
            {
                case "NetOffice": MsExcel_JsonWriter.WriteNetOffice(ReqWb(workbook), outputFilePath, log); break;
                case "Interop":   MsExcel_JsonWriter.WriteInterop(ReqWb(workbook), outputFilePath, log); break;
                case "Dynamic":   MsExcel_JsonWriter.WriteDynamic(ReqWb(workbook), outputFilePath, log); break;
                case "OpenXml":   MsExcel_JsonWriter.WriteOpenXmlFromRunningApp(ReqApp(application), outputFilePath, log); break;
                case "Direct":    MsOfficeDirectVbaRunner.RunWriteMacro(
                                      ReqApp(application).UnderlyingObject!, MsOfficeDirectVbaRunner.ExcelMacro, outputFilePath, log); break;
                default:          log($"Unknown method: {method}"); break;
            }

            static Excel.Workbook ReqWb(object? o) =>
                o as Excel.Workbook ?? throw new InvalidOperationException("No active workbook.");
            static Excel.Application ReqApp(object? o) =>
                o as Excel.Application ?? throw new InvalidOperationException("No Excel connection.");
        }

        public void WritePowerPointJson(string method, object? presentation, object? application, string outputFilePath, Action<string> log)
        {
            switch (method)
            {
                case "NetOffice": MsPowerPoint_JsonWriter.WriteNetOffice(ReqPrs(presentation), outputFilePath, log); break;
                case "Interop":   MsPowerPoint_JsonWriter.WriteInterop(ReqPrs(presentation), outputFilePath, log); break;
                case "Dynamic":   MsPowerPoint_JsonWriter.WriteDynamic(ReqPrs(presentation), outputFilePath, log); break;
                case "OpenXml":   MsPowerPoint_JsonWriter.WriteOpenXmlFromRunningApp(ReqApp(application), outputFilePath, log); break;
                case "Direct":    MsOfficeDirectVbaRunner.RunWriteMacro(
                                      ReqApp(application).UnderlyingObject!, MsOfficeDirectVbaRunner.PowerPointMacro, outputFilePath, log); break;
                default:          log($"Unknown method: {method}"); break;
            }

            static PowerPoint.Presentation ReqPrs(object? o) =>
                o as PowerPoint.Presentation ?? throw new InvalidOperationException("No active presentation.");
            static PowerPoint.Application ReqApp(object? o) =>
                o as PowerPoint.Application ?? throw new InvalidOperationException("No PowerPoint connection.");
        }
    }
}
