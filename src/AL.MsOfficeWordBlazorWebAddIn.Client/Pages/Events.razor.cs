using System.Runtime.Versioning;

using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class Events : ComponentBase, IAsyncDisposable
{
    [Inject]
    private WordEvents__Interop__Interface EventsInterop { get; set; } = default!;

    private sealed class EventDescriptor(string key, string label, bool enabled = true)
    {
        public string Key { get; } = key;
        public string Label { get; } = label;
        public bool Enabled { get; set; } = enabled;
    }

    // "selectionChanged" uses the Office common API (DocumentSelectionChanged). The content
    // control deleted/entered/exited events are per-ContentControl instance (WordApi 1.5)
    // and only cover controls that exist when watching starts.
    private readonly List<EventDescriptor> _events =
    [
        new("selectionChanged",      "Selection Changed"),
        new("paragraphAdded",        "Paragraph Added"),
        new("paragraphChanged",      "Paragraph Changed"),
        new("paragraphDeleted",      "Paragraph Deleted",      enabled: false),
        new("contentControlAdded",   "Content Control Added",  enabled: false),
        new("contentControlDeleted", "Content Control Deleted", enabled: false),
        new("contentControlEntered", "Content Control Entered", enabled: false),
        new("contentControlExited",  "Content Control Exited",  enabled: false),
    ];

    private DotNetObjectReference<Events>? _selfRef;
    private readonly List<string> _log = [];
    private bool _isWatching;
    private string? _error;

    private async Task StartWatching()
    {
        _error = null;
        try
        {
            var enabledKeys = _events.Where(e => e.Enabled).Select(e => e.Key).ToArray();
            _selfRef = DotNetObjectReference.Create(this);
            await EventsInterop.StartWatchingAsync(_selfRef, enabledKeys);
            _isWatching = true;
        }
        catch (Exception ex)
        {
            _error = ex.Message;
        }
        StateHasChanged();
    }

    private async Task StopWatching()
    {
        _error = null;
        try
        {
            await EventsInterop.StopWatchingAsync();
            _isWatching = false;
        }
        catch (Exception ex)
        {
            _error = ex.Message;
        }
        StateHasChanged();
    }

    private void ClearLog()
    {
        _log.Clear();
        StateHasChanged();
    }

    [JSInvokable]
    public void OnEventLogged(string eventName, string timestamp)
    {
        _log.Insert(0, $"[{timestamp}] {eventName}");
        if (_log.Count > 200) _log.RemoveAt(_log.Count - 1);
        StateHasChanged();
    }

    public async ValueTask DisposeAsync()
    {
        if (_isWatching)
        {
            try { await EventsInterop.StopWatchingAsync(); } catch { }
        }
        _selfRef?.Dispose();
    }
}
