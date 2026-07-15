using System.IO;

namespace JBC.ExploreTheWorld.UnitTests.CL;

public class OneDriveLocalPath_Helper_Tests : IDisposable
{
    // A fake local OneDrive sync root with one synced presentation in it.
    private readonly string _syncRoot;
    private readonly string _localFile;

    public OneDriveLocalPath_Helper_Tests()
    {
        _syncRoot = Path.Combine(Path.GetTempPath(), "ETW_OneDrive_" + Path.GetRandomFileName());
        _localFile = Path.Combine(_syncRoot, "Presentations", "Talks", "HowToThinkInBlazor", "HowToThinkInBlazor.pptx");

        Directory.CreateDirectory(Path.GetDirectoryName(_localFile)!);
        File.WriteAllText(_localFile, "test");
    }

    public void Dispose()
    {
        try { Directory.Delete(_syncRoot, true); } catch { }
    }

    [Fact]
    public void TryGetLocalPath_PersonalOneDriveUrl_MapsToConsumerSyncRoot()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://d.docs.live.net/7068DD7DE4BF55DD/Presentations/Talks/HowToThinkInBlazor/HowToThinkInBlazor.pptx",
            name => name == "OneDriveConsumer" ? _syncRoot : null);

        result.Should().Be(_localFile);
    }

    [Fact]
    public void TryGetLocalPath_PersonalOneDriveUrl_FallsBackToOneDriveVariable()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://d.docs.live.net/7068DD7DE4BF55DD/Presentations/Talks/HowToThinkInBlazor/HowToThinkInBlazor.pptx",
            name => name == "OneDrive" ? _syncRoot : null);

        result.Should().Be(_localFile);
    }

    [Fact]
    public void TryGetLocalPath_BusinessOneDriveUrl_SkipsPersonalSegmentsAndDecodes()
    {
        // OneDrive for Business URLs insert /personal/{user}/Documents/ before the path.
        var localFile = Path.Combine(_syncRoot, "My Decks", "Deck 1.pptx");
        Directory.CreateDirectory(Path.GetDirectoryName(localFile)!);
        File.WriteAllText(localFile, "test");

        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://contoso-my.sharepoint.com/personal/john_contoso_com/Documents/My%20Decks/Deck%201.pptx",
            name => name == "OneDriveCommercial" ? _syncRoot : null);

        result.Should().Be(localFile);
    }

    [Fact]
    public void TryGetLocalPath_FileNotSyncedLocally_ReturnsNull()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://d.docs.live.net/7068DD7DE4BF55DD/Presentations/NotSynced/Missing.pptx",
            name => name == "OneDriveConsumer" ? _syncRoot : null);

        result.Should().BeNull();
    }

    [Fact]
    public void TryGetLocalPath_UnknownHost_ReturnsNull()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://example.com/Presentations/Deck.pptx",
            name => _syncRoot);

        result.Should().BeNull();
    }

    [Fact]
    public void TryGetLocalPath_LocalPath_ReturnsNull()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            @"C:\Docs\Deck.pptx",
            name => _syncRoot);

        result.Should().BeNull();
    }

    [Fact]
    public void TryGetLocalPath_NoSyncRootVariables_ReturnsNull()
    {
        var result = OneDriveLocalPath_Helper.TryGetLocalPath(
            "https://d.docs.live.net/7068DD7DE4BF55DD/Presentations/Talks/HowToThinkInBlazor/HowToThinkInBlazor.pptx",
            name => null);

        result.Should().BeNull();
    }
}
