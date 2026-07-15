using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.PowerPointApi.Tools;
using Office = NetOffice.OfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF0123456706")]
    /// <summary>
    /// Hosts <see cref="MsPowerPoint_Watcher_Form"/> as a reusable UserControl
    /// so it can be embedded in VSTO Task Panes via ICTPFactory.CreateCTP.
    /// Implements <see cref="ITaskPane"/> so NetOffice injects the running PowerPoint
    /// application after the pane is created, and <see cref="IObjectSafety"/> so the
    /// Office ActiveX host accepts the control as safe (otherwise CreateCTP fails silently).
    /// </summary>
    public class MsPowerPoint_Watcher_UserControl : UserControl, ITaskPane, IObjectSafety
#if !NETFRAMEWORK
        , ICustomQueryInterface
#endif
    {
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x1;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x2;
#if !NETFRAMEWORK
        // Hides IProvideClassInfo / IProvideClassInfo2 from the Office CTP host so that
        // .NET 10 UserControls can be hosted as VSTO task panes without aborting on
        // missing typelib metadata (HRESULT 0x80131165). net481 controls expose these fine.
        private static readonly Guid _iidIProvideClassInfo  = new Guid("B196B283-BAB4-101A-B69C-00AA00341D07");
        private static readonly Guid _iidIProvideClassInfo2 = new Guid("A6BC3AC0-DBAA-11CE-9DE3-00AA004BB851");
        CustomQueryInterfaceResult ICustomQueryInterface.GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            if (iid == _iidIProvideClassInfo || iid == _iidIProvideClassInfo2)
                return CustomQueryInterfaceResult.Failed;
            return CustomQueryInterfaceResult.NotHandled;
        }
#endif

        int IObjectSafety.GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            return 0;
        }

        int IObjectSafety.SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions) => 0;

        private readonly MsPowerPoint_Watcher_Form _form;

        public MsPowerPoint_Watcher_UserControl()
        {
            _form = new MsPowerPoint_Watcher_Form();
            _form.TopLevel = false;
            _form.FormBorderStyle = FormBorderStyle.None;
            _form.Dock = DockStyle.Fill;
            Controls.Add(_form);
            _form.Show();
        }

        /// <summary>
        /// Injects a running PowerPoint.Application into the embedded watcher form.
        /// Called by NetOffice via <see cref="ITaskPane.OnConnection"/> when hosted as a task pane,
        /// or manually by the add-in when hosted in a floating form.
        /// </summary>
        public void InjectApplication(NetOffice.PowerPointApi.Application app)
            => _form.InjectApplication(app);

        // ITaskPane — NetOffice lifecycle callbacks for the CTP-hosted instance.
        public void OnConnection(NetOffice.PowerPointApi.Application application, Office._CustomTaskPane parentPane, object[] customArguments)
        {
            try { if (application != null) InjectApplication(application); } catch { }
        }

        public void OnDisconnection() { }
        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position) { }
        public void OnVisibleStateChanged(bool visible) { }
    }
}
