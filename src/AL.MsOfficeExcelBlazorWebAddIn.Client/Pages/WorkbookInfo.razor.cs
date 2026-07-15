using System.Runtime.Versioning;

using JBC.ExploreTheWorld.DL.MsOfficeApi.MsOfficeJs.Excel_Impl;

using Microsoft.AspNetCore.Components;

namespace ExploreTheWorld.AL.MsOfficeExcelBlazorWebAddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class WorkbookInfo : ComponentBase
{
    [Inject]
    private ExcelWorkbookInfo__Interop__Interface WorkbookInfoInterop { get; set; } = default!;

    private ExcelWorkbookInfo_Row? _info;
    private bool _isLoading;
    private string? _error;

    private IQueryable<InfoRow>? InfoItems => _info is null ? null : new[]
    {
        new InfoRow("Workbook Name",  _info.WorkbookName),
        new InfoRow("Active Sheet",   _info.ActiveSheetName),
        new InfoRow("Sheet Count",    _info.SheetCount.ToString()),
        new InfoRow("Used Range",     _info.UsedRangeAddress),
    }.AsQueryable();

    private async Task RefreshWorkbookInfo()
    {
        _isLoading = true;
        _error = null;
        _info = null;
        StateHasChanged();

        try
        {
            _info = await WorkbookInfoInterop.GetWorkbookInfoAsync();
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
