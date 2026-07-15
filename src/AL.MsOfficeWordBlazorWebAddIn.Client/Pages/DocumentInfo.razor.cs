using System.Runtime.Versioning;

using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Word_Impl;

using Microsoft.AspNetCore.Components;

namespace ExploreTheWorld.AL.MsOfficeWordBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class DocumentInfo : ComponentBase
{
    [Inject]
    private WordDocumentInfo__Interop__Interface DocumentInfoInterop { get; set; } = default!;

    private WordDocumentInfo_Row? _info;
    private bool _isLoading;
    private string? _error;

    private IQueryable<InfoRow>? InfoItems => _info is null ? null : new[]
    {
        new InfoRow("Title",           _info.Title),
        new InfoRow("Author",          _info.Author),
        new InfoRow("Word Count",      _info.WordCount.ToString()),
        new InfoRow("Paragraph Count", _info.ParagraphCount.ToString()),
        new InfoRow("Page Count",      _info.PageCount.ToString()),
        new InfoRow("Revision",        _info.Revision.ToString()),
    }.AsQueryable();

    private async Task RefreshDocumentInfo()
    {
        _isLoading = true;
        _error = null;
        _info = null;
        StateHasChanged();

        try
        {
            _info = await DocumentInfoInterop.GetDocumentInfoAsync();
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
