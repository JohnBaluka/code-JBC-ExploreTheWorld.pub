using System;
using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// Host-supplied entry point to the "Save as JSON" writer for the WinForms watcher forms. The
    /// forms live in the AL WinForms UI libraries, which reference no DL repo `_Impl` project, so
    /// they cannot instantiate the writer themselves. Each host (WinFormApp, the VSTO add-ins, the
    /// MAUI WinUI head) assigns <see cref="Current"/> once at startup to the concrete
    /// <c>MsOfficeSaveAsJsonWriter</c> from the <c>DL.MsOfficeApi_Impl</c> composition
    /// project. A static seam (rather than DI) is used because the net481 watcher forms are
    /// constructed without a service provider (VSTO declarative panes / parameterless forms).
    /// </summary>
    public static class MsOfficeSaveAsJsonWriterProvider
    {
        public static MsOfficeSaveAsJsonWriter__Interface? Current { get; set; }

        /// <summary>Returns the configured writer, or throws a clear error if a host forgot to set it.</summary>
        public static MsOfficeSaveAsJsonWriter__Interface Require() =>
            Current ?? throw new InvalidOperationException(
                "No MsOfficeSaveAsJsonWriter has been configured. The host application must set " +
                "MsOfficeSaveAsJsonWriterProvider.Current at startup (see DL.MsOfficeApi_Impl).");
    }
}
