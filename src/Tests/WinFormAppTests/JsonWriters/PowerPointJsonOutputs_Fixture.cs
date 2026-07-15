using JBC.ExploreTheWorld.DL.MsOfficeApi;
using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice;
using DynamicRepos = JBC.ExploreTheWorld.DL.MsOfficeApi.Dynamic_Impl;
using NetOfficeRepos = JBC.ExploreTheWorld.DL.MsOfficeApi.NetOffice_Impl;
using OpenXmlRepos = JBC.ExploreTheWorld.DL.MsOfficeApi.OpenXml_Impl;
using PP = JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint;

namespace JBC.ExploreTheWorld.WinFormAppTests.JsonWriters;

/// <summary>
/// Generates one small presentation and writes its canonical JSON once with each
/// "Save as JSON" method the WinFormApp watcher offers (NetOffice, Dynamic, OpenXml),
/// so the comparison tests can assert the outputs against each other. The NetOffice
/// and Dynamic writers automate a hidden PowerPoint instance; when PowerPoint is not
/// installed <see cref="PowerPointAvailable"/> is false and those outputs stay null.
/// </summary>
public sealed class PowerPointJsonOutputs_Fixture : IAsyncLifetime
{
    private static readonly List<MsOfficeCountry_Row> Countries = new()
    {
        new MsOfficeCountry_Row("Australia", "AU", "AUS"),
        new MsOfficeCountry_Row("Brazil",    "BR", "BRA"),
    };

    private readonly List<string> _tempFiles = new();

    public bool PowerPointAvailable { get; private set; }
    public string PptxPath { get; private set; } = string.Empty;
    public string NetOfficeJsonPath { get; private set; } = string.Empty;
    public string DynamicJsonPath { get; private set; } = string.Empty;
    public string OpenXmlJsonPath { get; private set; } = string.Empty;
    public PP.Presentation? NetOfficePresentation { get; private set; }
    public PP.Presentation? DynamicPresentation { get; private set; }
    public PP.Presentation? OpenXmlPresentation { get; private set; }

    public async Task InitializeAsync()
    {
        PowerPointAvailable = Type.GetTypeFromProgID("PowerPoint.Application") != null;

        PptxPath = NewTempPath(".pptx");
        NetOfficeJsonPath = NewTempPath(".netoffice.pptx.json");
        DynamicJsonPath = NewTempPath(".dynamic.pptx.json");
        OpenXmlJsonPath = NewTempPath(".openxml.pptx.json");

        var openXmlRepo = new OpenXmlRepos.MsPowerPoint_OpenXml__Repo();
        await openXmlRepo.ExportAsync(Countries, PptxPath, _ => { });

        await openXmlRepo.WriteDocumentJsonAsync(PptxPath, OpenXmlJsonPath, _ => { });
        OpenXmlPresentation = MsOfficeJsonSerializer.ReadFromFile<PP.Presentation>(OpenXmlJsonPath);

        if (!PowerPointAvailable) return;

        await new NetOfficeRepos.MsPowerPoint_NetOffice__Repo()
            .WriteDocumentJsonAsync(PptxPath, NetOfficeJsonPath, _ => { });
        NetOfficePresentation = MsOfficeJsonSerializer.ReadFromFile<PP.Presentation>(NetOfficeJsonPath);

        await new DynamicRepos.MsPowerPoint_Dynamic__Repo()
            .WriteDocumentJsonAsync(PptxPath, DynamicJsonPath, _ => { });
        DynamicPresentation = MsOfficeJsonSerializer.ReadFromFile<PP.Presentation>(DynamicJsonPath);
    }

    public Task DisposeAsync()
    {
        foreach (var path in _tempFiles)
        {
            try { File.Delete(path); } catch { }
        }

        return Task.CompletedTask;
    }

    private string NewTempPath(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), "ETW_" + Guid.NewGuid().ToString("N") + extension);
        _tempFiles.Add(path);
        return path;
    }
}
