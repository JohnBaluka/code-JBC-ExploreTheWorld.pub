using System.IO;

namespace JBC.ExploreTheWorld.UnitTests.CL;

public class SaveAsJson_Helper_Tests
{
    private static readonly string Documents =
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    [Fact]
    public void BuildDefaultPath_NullCurrentFilePath_UsesFallbackNameInDocuments()
    {
        var result = SaveAsJson_Helper.BuildDefaultPath(null, "ETW_CountriesNow");

        result.Should().Be(Path.Combine(Documents, "ETW_CountriesNow.json"));
    }

    [Fact]
    public void BuildDefaultPath_WhitespaceCurrentFilePath_UsesFallbackNameInDocuments()
    {
        var result = SaveAsJson_Helper.BuildDefaultPath("   ", "ETW_CountriesNow");

        result.Should().Be(Path.Combine(Documents, "ETW_CountriesNow.json"));
    }

    [Fact]
    public void BuildDefaultPath_FullFilePath_AppendsJsonInSameDirectory()
    {
        var result = SaveAsJson_Helper.BuildDefaultPath(@"C:\Docs\ETW_CountriesNow.docx", "fallback");

        result.Should().Be(@"C:\Docs\ETW_CountriesNow.docx.json");
    }

    [Fact]
    public void BuildDefaultPath_BareDocumentName_FallsBackToDocumentsFolder()
    {
        // An unsaved document (e.g. "Document1") has no directory component.
        var result = SaveAsJson_Helper.BuildDefaultPath("Document1", "fallback");

        result.Should().Be(Path.Combine(Documents, "Document1.json"));
    }

    [Fact]
    public void BuildDefaultPath_UnmappableCloudUrl_FallsBackToDocumentsWithFileName()
    {
        // A cloud URL that has no local OneDrive mapping keeps the document's file name
        // but defaults to the Documents folder instead of the https path.
        var result = SaveAsJson_Helper.BuildDefaultPath(
            "https://d.docs.live.net/0000000000000000/Presentations/NoSuchLocalCopy/Deck%201.pptx", "fallback");

        result.Should().Be(Path.Combine(Documents, "Deck 1.pptx.json"));
    }

    [Fact]
    public void BuildDefaultPath_MappableCloudUrl_UsesLocalOneDrivePath()
    {
        // With a real local OneDrive sync root, a cloud URL maps to the synced file's folder.
        var oneDriveRoot = Environment.GetEnvironmentVariable("OneDriveConsumer")
            ?? Environment.GetEnvironmentVariable("OneDrive");
        if (string.IsNullOrEmpty(oneDriveRoot) || !Directory.Exists(oneDriveRoot))
            return; // no OneDrive sync client on this machine

        var localFile = Path.Combine(oneDriveRoot, "ETW_Test_" + Path.GetRandomFileName() + ".pptx");
        File.WriteAllText(localFile, "test");
        try
        {
            var url = "https://d.docs.live.net/7068DD7DE4BF55DD/" + Path.GetFileName(localFile);

            var result = SaveAsJson_Helper.BuildDefaultPath(url, "fallback");

            result.Should().Be(localFile + ".json");
        }
        finally
        {
            File.Delete(localFile);
        }
    }
}
