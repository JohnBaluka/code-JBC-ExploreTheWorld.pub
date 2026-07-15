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
    public partial class MsPowerPoint_Watcher__Component : Base__RadzenComponent
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
                _lastActiveFile = Watcher_AppService.PowerPointActiveFile;
                _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                    _lastActiveFile, "ETW_MsPowerPoint", _jsonMethod);

                if (Watcher_AppService.IsOfficeAddinHost && _jsonMethod == "OpenXml")
                    _jsonMethod = "NetOffice";

                _selectedFile = Watcher_AppService.PowerPointActiveFile;
                Watcher_AppService.PowerPointStateChanged += OnPowerPointStateChanged;
            }
            catch (Exception ex)
            {
                Watcher_AppService.PowerPoint_AppendLog($"✘ Init error: {ex.Message}");
            }
        }

        private void OnPowerPointStateChanged()
        {
            _selectedFile = Watcher_AppService.PowerPointActiveFile;

            // When the active presentation changes, refresh the default .json path
            // to follow the standard: <current file name>.json.
            if (Watcher_AppService.PowerPointActiveFile != _lastActiveFile)
            {
                _lastActiveFile = Watcher_AppService.PowerPointActiveFile;
                _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                    _lastActiveFile, "ETW_MsPowerPoint", _jsonMethod);
            }

            InvokeAsync(StateHasChanged);
        }

        private void OnJsonMethodChanged(object value)
        {
            if (value is string method)
                _jsonMethod = method;

            // The method is embedded in the default file name (e.g. filename-NetOffice.pptx.json),
            // so recompute the default .json path whenever the method changes.
            _outputFilePath = SaveAsJson_Helper.BuildDefaultPath(
                _lastActiveFile, "ETW_MsPowerPoint", _jsonMethod);
        }

        private void OnActiveDocChanged(object value)
        {
            if (value is string prsName && !string.IsNullOrEmpty(prsName))
                Watcher_AppService.PowerPointActivateDocumentAction?.Invoke(prsName);
        }

        private async Task SaveAsJsonAsync()
        {
            if (string.IsNullOrWhiteSpace(_outputFilePath))
            {
                Watcher_AppService.PowerPoint_AppendLog("ERROR: No output file path specified.");
                return;
            }

            if (Watcher_AppService.PowerPointSaveAsJsonFunc != null)
            {
                Watcher_AppService.PowerPoint_AppendLog($"Save As Json ({_jsonMethod}): {_outputFilePath}");
                try
                {
                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                    await Watcher_AppService.PowerPointSaveAsJsonFunc(_jsonMethod, _outputFilePath,
                        msg => Watcher_AppService.PowerPoint_AppendLog(msg));
                    stopwatch.Stop();
                    Watcher_AppService.PowerPoint_AppendLog($"✔ Saved: {_outputFilePath} ({Duration_Helper.Format(stopwatch.Elapsed)})");
                }
                catch (Exception ex)
                {
                    Watcher_AppService.PowerPoint_AppendLog($"✘ Save failed: {ex.Message}");
                }
            }
            else
            {
                Watcher_AppService.PowerPoint_AppendLog($"Saving watcher state to: {_outputFilePath}");
                try
                {
                    var data = new
                    {
                        Source     = "WatcherEvent_AppService",
                        Timestamp  = DateTime.Now,
                        OpenFiles  = Watcher_AppService.PowerPointOpenFiles,
                        ActiveFile = Watcher_AppService.PowerPointActiveFile,
                        EventLog   = Watcher_AppService.PowerPointLog
                    };
                    var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                    var dir  = Path.GetDirectoryName(_outputFilePath);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    await File.WriteAllTextAsync(_outputFilePath, json, Encoding.UTF8);
                    Watcher_AppService.PowerPoint_AppendLog($"✔ Saved: {_outputFilePath}");
                }
                catch (Exception ex)
                {
                    Watcher_AppService.PowerPoint_AppendLog($"✘ Save failed: {ex.Message}");
                }
            }
        }

        private void ToggleEventLog(WatcherEventToggle toggle)
        {
            Watcher_AppService.PowerPoint_SetEventLog(toggle.Name, !toggle.Log);
        }

        private void OnConnectClick()
        {
            if (Watcher_AppService.PowerPointConnected)
                Watcher_AppService.PowerPointDisconnectAction?.Invoke();
            else
                Watcher_AppService.PowerPointConnectAction?.Invoke();
        }

        private void ClearLog() => Watcher_AppService.PowerPoint_ClearLog();

        public override void Dispose()
        {
            Watcher_AppService.PowerPointStateChanged -= OnPowerPointStateChanged;
            base.Dispose();
        }
    }
}
