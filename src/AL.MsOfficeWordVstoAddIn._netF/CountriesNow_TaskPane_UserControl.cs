using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using JBC.ExploreTheWorld.AL.WinFormsLib;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceData.SqliteDb_Impl;
using NetOffice.WordApi.Tools;
using Office = NetOffice.OfficeApi;

namespace JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn
{
    // Wrapper that gives CountriesNowSpace_UserControl a parameterless constructor so that
    // ICTPFactory.CreateCTP can instantiate it without dependency-injection infrastructure.
    [ComVisible(true)]
    [Guid("C1D2E3F4-A5B6-7890-ABCD-EF0123456710")]
    [ProgId("JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn.CountriesNow_TaskPane_UserControl")]
    public partial class CountriesNow_TaskPane_UserControl : UserControl, ITaskPane, IObjectSafety
    {
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x1;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x2;

        // IObjectSafety — the Office ActiveX host queries this before hosting the pane content;
        // without it, ICTPFactory.CreateCTP fails silently and the pane never appears.
        int IObjectSafety.GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return 0;
        }

        int IObjectSafety.SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions) => 0;

        public CountriesNow_TaskPane_UserControl()
        {
            InitializeComponent();
            try
            {
                var dbPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JBC", "ExploreTheWorld", "etw.db");
                var dbFactory = ExploreTheWorldSqliteDb.CreateFactory($"Data Source={dbPath}");
                try { JBC.ExploreTheWorld.DL.CountriesNowSpaceData.ServiceCollectionExtensions.EnsureExploreTheWorldDbCreated(dbFactory); } catch { }
                var api     = new CountriesNowSpaceApi__Repo();
                var dbMgr   = new CountriesNowSpaceApiManager__Repo(dbFactory);
                var manager = new CountriesNowSpaceManager__Service(api, dbMgr);
                var inner   = new CountriesNowSpace_UserControl(manager) { Dock = DockStyle.Fill };
                inner.EnableAddinExportMode("Word");
                Controls.Add(inner);
            }
            catch (Exception ex)
            {
                Controls.Add(new Label
                {
                    Text     = $"Error initialising Countries API: {ex.Message}",
                    Dock     = DockStyle.Fill,
                    AutoSize = false
                });
            }
        }

        public void OnConnection(
            NetOffice.WordApi.Application application,
            Office._CustomTaskPane parentPane,
            object[] customArguments) { }

        public void OnDisconnection() { }
        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position) { }
        public void OnVisibleStateChanged(bool visible) { }
    }
}
