using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    public partial class ExportLog_Form : Form
    {
        private readonly ExportType _exportType;
        private readonly ExportMethod _exportMethod;
        private readonly IList<MsOfficeCountry_Row> _countries;
        private readonly MsOfficeExportManager__Service _exportManager;

        public ExportLog_Form(ExportType exportType, ExportMethod exportMethod,
            IList<MsOfficeCountry_Row> countries, MsOfficeExportManager__Service exportManager)
        {
            InitializeComponent();
            _exportType    = exportType;
            _exportMethod  = exportMethod;
            _countries     = countries;
            _exportManager = exportManager;
            Text = $"{exportType} Export — {GetMethodDisplayName(exportMethod)}";
        }

        private static string GetMethodDisplayName(ExportMethod method) => method switch
        {
            ExportMethod.Interop   => "Using Microsoft.Office.Interop",
            ExportMethod.NetOffice => "Using NetOffice",
            ExportMethod.OpenXml   => "Using OpenXML",
            ExportMethod.Dynamic   => "Using dynamic COM automation",
            _                      => method.ToString()
        };

        private static string GetFileFilter(ExportType type) => type switch
        {
            ExportType.Word        => "Word Document (*.docx)|*.docx",
            ExportType.Excel       => "Excel Workbook (*.xlsx)|*.xlsx",
            ExportType.PowerPoint  => "PowerPoint Presentation (*.pptx)|*.pptx",
            _                      => "All Files (*.*)|*.*"
        };

        private static string GetFileExtension(ExportType type) => type switch
        {
            ExportType.Word        => "docx",
            ExportType.Excel       => "xlsx",
            ExportType.PowerPoint  => "pptx",
            _                      => "dat"
        };

        private void AppendLog(string message)
        {
            if (IsHandleCreated)
            {
                if (InvokeRequired)
                {
                    Invoke(() => AppendLog(message));
                    return;
                }
                rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] {message}{Environment.NewLine}");
                rtbLog.ScrollToCaret();
            }
        }

        private async void btnRunExport_Click(object sender, EventArgs e)
        {
            btnRunExport.Enabled = false;

            using var dlg = new SaveFileDialog
            {
                Title      = $"Save {_exportType} Document",
                Filter     = GetFileFilter(_exportType),
                DefaultExt = GetFileExtension(_exportType),
                FileName   = JBC.ExploreTheWorld.CL.MsOfficeExportName_Helper.BuildStem(
                    "ETW_CountriesNow", _exportMethod.ToString(), fromAccessDb: false)
            };

            if (dlg.ShowDialog(this) != DialogResult.OK)
            {
                btnRunExport.Enabled = true;
                return;
            }

            var filePath = dlg.FileName;
            AppendLog($"Export target : {filePath}");
            AppendLog($"Countries     : {_countries.Count}");
            AppendLog("─────────────────────────────────────────────────────");

            try
            {
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                await RunExportAsync(filePath);
                stopwatch.Stop();
                AppendLog("─────────────────────────────────────────────────────");
                AppendLog($"✔ Export completed successfully in {Duration_Helper.Format(stopwatch.Elapsed)}.");
            }
            catch (Exception ex)
            {
                AppendLog("─────────────────────────────────────────────────────");
                AppendLog($"✘ Export failed: {ex.Message}");
                MessageBox.Show(ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnRunExport.Enabled = true;
            }
        }

        private Task RunExportAsync(string filePath) =>
            _exportManager.ExportAsync(
                _countries, _exportType.ToString(), _exportMethod.ToString(), filePath, AppendLog);
    }
}
