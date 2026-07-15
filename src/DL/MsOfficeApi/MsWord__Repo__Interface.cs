using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi
{
    public interface MsWord__Repo__Interface
    {
        Task ExportAsync(IList<MsOfficeCountry_Row> countries, string filePath, Action<string> log);
        Task WriteDocumentJsonAsync(string sourcePath, string outputJsonPath, Action<string> log);
    }
}
