using System;
using System.IO;
using NetOffice.PowerPointApi;
using DynamicWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl.JsonWriters;
using InteropWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.Interop_Impl.JsonWriters;
using NetOfficeWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl.JsonWriters;
using OpenXmlWriters = JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl.JsonWriters;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi_Impl
{
    // Thin wrapper over the DL "Save as JSON" writers so the watcher forms keep their
    // existing call sites. Both paths emit the canonical
    // JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint entity schema (same as the VBA writers).
    internal static class MsPowerPoint_JsonWriter
    {
        /// <summary>Writes a full JSON representation of the presentation using the NetOffice COM object model.</summary>
        internal static void WriteNetOffice(Presentation prs, string filePath, Action<string> log)
        {
            if (prs == null) throw new ArgumentNullException(nameof(prs));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            NetOfficeWriters.MsPowerPointJsonWriter.WritePresentationToJsonFile(prs, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the presentation using the late-bound dynamic COM writers (DL.MsOfficeApi.Dynamic_Impl).</summary>
        internal static void WriteDynamic(Presentation prs, string filePath, Action<string> log)
        {
            if (prs == null) throw new ArgumentNullException(nameof(prs));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            DynamicWriters.MsPowerPointJsonWriter.WritePresentationComToJsonFile(prs.UnderlyingObject, filePath, null, log);
        }

        /// <summary>Writes a full JSON representation of the presentation using the strongly-typed Microsoft.Office.Interop COM object model (DL.MsOfficeApi.Interop_Impl).</summary>
        internal static void WriteInterop(Presentation prs, string filePath, Action<string> log)
        {
            if (prs == null) throw new ArgumentNullException(nameof(prs));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            // The watcher holds a NetOffice wrapper; unwrap to the raw COM object and cast to the
            // Interop interface (a QueryInterface on the same live PowerPoint.Presentation COM object).
            var interopPresentation = (Microsoft.Office.Interop.PowerPoint.Presentation)prs.UnderlyingObject;
            InteropWriters.MsPowerPointJsonWriter.WritePresentationToJsonFile(interopPresentation, filePath, null, log);
        }

        /// <summary>Writes a JSON representation of the .pptx file at <paramref name="sourcePptxPath"/> using OpenXML (no COM required).</summary>
        internal static void WriteOpenXml(string sourcePptxPath, string outputJsonFilePath, Action<string> log)
        {
            if (string.IsNullOrEmpty(sourcePptxPath)) throw new ArgumentException("sourcePptxPath is required.", nameof(sourcePptxPath));
            if (string.IsNullOrEmpty(outputJsonFilePath)) throw new ArgumentException("outputJsonFilePath is required.", nameof(outputJsonFilePath));

            OpenXmlWriters.MsPowerPointJsonWriter.WritePresentationFileToJson(sourcePptxPath, outputJsonFilePath, null, log);
        }

        /// <summary>
        /// Writes the running app's active presentation to JSON using OpenXML. OpenXML reads the
        /// file on disk while PowerPoint holds a lock on it, so the presentation is saved, closed,
        /// read, and reopened. Cloud (AutoSave/OneDrive) presentations have an https path OpenXML
        /// cannot read; those are read from a temporary local copy instead (no close/reopen).
        /// </summary>
        internal static void WriteOpenXmlFromRunningApp(NetOffice.PowerPointApi.Application pptApp, string outputJsonFilePath, Action<string> log)
        {
            if (pptApp == null) throw new ArgumentNullException(nameof(pptApp));

            var prs = pptApp.ActivePresentation
                ?? throw new InvalidOperationException("No active presentation.");

            string documentFolder = string.Empty;
            try { documentFolder = prs.Path; } catch { }
            if (string.IsNullOrEmpty(documentFolder))
                throw new InvalidOperationException(
                    "The active presentation has never been saved — save it first, or use the NetOffice method.");

            string fullName = prs.FullName;

            if (fullName.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                log("Cloud (AutoSave) presentation detected — reading a temporary local copy...");
                string extension = Path.GetExtension(fullName);
                if (string.IsNullOrEmpty(extension)) extension = ".pptx";
                string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + extension);

                prs.SaveCopyAs(tempPath);
                try
                {
                    WriteOpenXml(tempPath, outputJsonFilePath, log);
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }
                return;
            }

            log("Closing active presentation for OpenXML read...");
            try { prs.Save(); } catch { }
            prs.Close();
            try
            {
                WriteOpenXml(fullName, outputJsonFilePath, log);
            }
            finally
            {
                log("Reopening presentation...");
                pptApp.Presentations.Open(fullName);
            }
        }
    }
}
