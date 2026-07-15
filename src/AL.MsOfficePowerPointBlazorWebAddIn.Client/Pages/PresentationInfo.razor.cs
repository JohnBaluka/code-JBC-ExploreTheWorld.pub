using System.Runtime.Versioning;

using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.PowerPoint_Impl;

using Microsoft.AspNetCore.Components;

namespace ExploreTheWorld.AL.MsOfficePowerPointBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class PresentationInfo : ComponentBase
{
    [Inject]
    private PowerPointPresentationInfo__Interop__Interface PresentationInfoInterop { get; set; } = default!;

    private PowerPointPresentationInfo_Row? _info;
    private bool _isLoading;
    private string? _error;

    private IQueryable<InfoRow>? InfoItems => _info is null ? null : new[]
    {
        new InfoRow("Title",        _info.Title),
        new InfoRow("Slide Count",  _info.SlideCount.ToString()),
        new InfoRow("Slide Width",  _info.SlideWidth.ToString("F2")),
        new InfoRow("Slide Height", _info.SlideHeight.ToString("F2")),
        new InfoRow("Author",       _info.Author),
    }.AsQueryable();

    private async Task RefreshPresentationInfo()
    {
        _isLoading = true;
        _error = null;
        _info = null;
        StateHasChanged();

        try
        {
            _info = await PresentationInfoInterop.GetPresentationInfoAsync();
            if (!string.IsNullOrEmpty(_info.ErrorMessage))
            {
                _error = _info.ErrorMessage;
                _info = null;
            }
        }
        catch (Exception ex)
        {
            _error = ex.Message;
        }
        finally
        {
            _isLoading = false;
            StateHasChanged();
        }
    }
}
