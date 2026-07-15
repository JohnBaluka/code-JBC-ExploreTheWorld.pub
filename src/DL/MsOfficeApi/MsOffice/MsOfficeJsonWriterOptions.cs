using System;
using System.IO;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice
{
    // Options shared by all "Save as JSON" writers (VBA mirrors these, NetOffice, OpenXML, Office.js interop).
    public class MsOfficeJsonWriterOptions
    {
        // How image/blob properties are written. Default is base64 embedded in the JSON.
        public BlobOutput_Enum BlobOutput { get; set; } = BlobOutput_Enum.Base64;

        // Folder where blob files are written when BlobOutput = SeparateFiles.
        // When null, "{jsonFileName}_Files" beside the output .json file is used.
        public string? BlobFolderPath { get; set; }

        // Also saves the original (unconverted) media blobs to an "images" sub-folder
        // of the blob folder, one file per unique image part. Only applies when the
        // writer outputs to a real file path (not for in-browser scenarios).
        public bool SaveOriginalMedia { get; set; } = true;

        public string GetBlobFolderPath(string outputJsonFilePath)
        {
            if (!string.IsNullOrEmpty(BlobFolderPath)) return BlobFolderPath!;

            string folder = Path.GetDirectoryName(outputJsonFilePath) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(outputJsonFilePath) + "_Files";

            return Path.Combine(folder, name);
        }
    }
}
