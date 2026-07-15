using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.AL.BlazorLib;
using JBC.ExploreTheWorld.CL;
using Microsoft.AspNetCore.Components;

namespace JBC.ExploreTheWorld.AL.BlazorLib.Watcher
{
    public partial class MsWord_Watcher__Component : Base__RadzenComponent
    {
        [Inject] protected WatcherEvent_AppService Watcher_AppService { get; set; } = default!;

        private string? _selectedFile;
        private string? _lastActiveFile;
        private string  _outputFilePath = string.Empty;
        private string  _jsonMethod     = "NetOffice";

        // OpenXml is excluded in Office add-in hosts (the active document must stay open).
        private string[] _jsonMethods =>
            Watcher_AppService.IsOfficeAddinHost
                ? ["Direct", "Interop", "Dynamic", "NetOffice"]
                : ["Direct", "Interop", "Dynamic", "NetOffice", "OpenXml"];

        protected override void OnInitialized()
        {
            try
            {
                _lastActiveFile = Watcher_AppService.WordActiveFile;
                _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                    _lastActiveFile, "ETW_MsWord", _jsonMethod);

                if (Watcher_AppService.IsOfficeAddinHost && _jsonMethod == "OpenXml")
                    _jsonMethod = "NetOffice";

                _selectedFile = Watcher_AppService.WordActiveFile;
                Watcher_AppService.WordStateChanged += OnWordStateChanged;
            }
            catch (Exception ex)
            {
                Watcher_AppService.Word_AppendLog($"✘ Init error: {ex.Message}");
            }
        }

        private void OnWordStateChanged()
        {
            _selectedFile = Watcher_AppService.WordActiveFile;

            // When the active document changes, refresh the default .json path
            // to follow the standard: <current file name>.json.
            if (Watcher_AppService.WordActiveFile != _lastActiveFile)
            {
                _lastActiveFile = Watcher_AppService.WordActiveFile;
                _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                    _lastActiveFile, "ETW_MsWord", _jsonMethod);
            }

            InvokeAsync(StateHasChanged);
        }

        private void OnJsonMethodChanged(object value)
        {
            if (value is string method)
                _jsonMethod = method;

            // The method is embedded in the default file name (e.g. filename-NetOffice.docx.json),
            // so recompute the default .json path whenever the method changes.
            _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                _lastActiveFile, "ETW_MsWord", _jsonMethod);
        }

        private void OnActiveDocChanged(object value)
        {
            if (value is string docName && !string.IsNullOrEmpty(docName))
                Watcher_AppService.WordActivateDocumentAction?.Invoke(docName);
        }

        private async Task SaveAsJsonAsync()
        {
            if (string.IsNullOrWhiteSpace(_outputFilePath))
            {
                Watcher_AppService.Word_AppendLog("ERROR: No output file path specified.");
                return;
            }

            if (Watcher_AppService.WordSaveAsJsonFunc != null)
            {
                Watcher_AppService.Word_AppendLog($"Save As Json ({_jsonMethod}): {_outputFilePath}");
                try
                {
                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                    await Watcher_AppService.WordSaveAsJsonFunc(_jsonMethod, _outputFilePath,
                        msg => Watcher_AppService.Word_AppendLog(msg));
                    stopwatch.Stop();
                    Watcher_AppService.Word_AppendLog($"✔ Saved: {_outputFilePath} ({Duration_Helper.Format(stopwatch.Elapsed)})");
                }
                catch (Exception ex)
                {
                    Watcher_AppService.Word_AppendLog($"✘ Save failed: {ex.Message}");
                }
            }
            else
            {
                Watcher_AppService.Word_AppendLog($"Saving watcher state to: {_outputFilePath}");
                try
                {
                    var data = new
                    {
                        Source     = "WatcherEvent_AppService",
                        Timestamp  = DateTime.Now,
                        OpenFiles  = Watcher_AppService.WordOpenFiles,
                        ActiveFile = Watcher_AppService.WordActiveFile,
                        EventLog   = Watcher_AppService.WordLog
                    };
                    var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                    var dir  = Path.GetDirectoryName(_outputFilePath);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    await File.WriteAllTextAsync(_outputFilePath, json, Encoding.UTF8);
                    Watcher_AppService.Word_AppendLog($"✔ Saved: {_outputFilePath}");
                }
                catch (Exception ex)
                {
                    Watcher_AppService.Word_AppendLog($"✘ Save failed: {ex.Message}");
                }
            }
        }

        private void ToggleEventLog(WatcherEventToggle toggle)
        {
            Watcher_AppService.Word_SetEventLog(toggle.Name, !toggle.Log);
        }

        private void OnConnectClick()
        {
            if (Watcher_AppService.WordConnected)
                Watcher_AppService.WordDisconnectAction?.Invoke();
            else
                Watcher_AppService.WordConnectAction?.Invoke();
        }

        private void ClearLog() => Watcher_AppService.Word_ClearLog();

        public override void Dispose()
        {
            Watcher_AppService.WordStateChanged -= OnWordStateChanged;
            base.Dispose();
        }
    }
}
