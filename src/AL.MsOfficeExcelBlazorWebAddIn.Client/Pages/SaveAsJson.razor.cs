using System.Runtime.Versioning;

using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class SaveAsJson : ComponentBase
{
    [Inject]
    private ExcelSaveAsJson__Interop__Interface SaveAsJsonInterop { get; set; } = default!;

    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    private string? _json;
    private string? _error;
    private string? _fileName;
    private bool _isLoading;

    private async Task BuildJsonAsync()
    {
        _isLoading = true;
        _error     = null;
        _json      = null;
        StateHasChanged();
        try
        {
            var result = await SaveAsJsonInterop.GetWorkbookAsJsonAsync();
            _fileName = result.FileName;
            if (!string.IsNullOrEmpty(result.Error))
                _error = result.Error;
            else
                _json = result.Json;
        }
        catch (Exception ex) { _error = ex.Message; }
        _isLoading = false;
        StateHasChanged();
    }

    private async Task CopyToClipboardAsync()
    {
        if (_json is null) return;
        try { await JSRuntime.InvokeVoidAsync("navigator.clipboard.writeText", _json); }
        catch { }
    }

    private async Task DownloadAsync()
    {
        if (_json is null) return;
        var name = string.IsNullOrWhiteSpace(_fileName) ? "MsExcel.json" : _fileName + ".json";
        try { await SaveAsJsonInterop.DownloadJsonAsync(_json, name); }
        catch { }
    }
}
