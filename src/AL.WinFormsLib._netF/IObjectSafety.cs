using System;
using System.Runtime.InteropServices;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// Classic ActiveX <c>IObjectSafety</c> COM interface (IID CB5BDC81-93C1-11CF-8F20-00805F2CD064).
    /// Office hosts a Custom Task Pane's content control through its ActiveX container, which queries
    /// this interface to decide whether the control is safe to initialize/script. A .NET
    /// <see cref="System.Windows.Forms.UserControl"/> that is registered by hand (no full regasm
    /// component-category registration) is treated as untrusted, and <c>ICTPFactory.CreateCTP</c>
    /// fails silently — the pane is never created. Implementing this interface (returning
    /// "safe for initializing + scripting") lets the net481 VSTO/NetOffice add-ins host their panes.
    /// Not part of the BCL, so it is declared here and shared by every net481 CTP-hosted control.
    /// </summary>
    [ComImport]
    [Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions);

        [PreserveSig]
        int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions);
    }
}
