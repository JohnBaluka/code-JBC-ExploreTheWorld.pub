using JBC.ExploreTheWorld.DL.MsOfficeApi;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// Host-supplied Office export repo factory for the WinForms UI libraries. Some WinForms UI
    /// items (e.g. the net481 <c>CountriesNowSpace_UserControl</c> hosted as a VSTO task pane) build
    /// their own <c>MsOfficeExportManager__Service</c> and therefore need the platform-specific
    /// export factory, which lives in the <c>DL.MsOfficeApi_Impl</c> composition project the
    /// UI libraries do not reference. Each host assigns <see cref="Current"/> once at startup to the
    /// concrete <c>MsOfficeExportRepoFactory</c>. See <c>MsOfficeSaveAsJsonWriterProvider</c> for the
    /// analogous Save-As-JSON seam.
    /// </summary>
    public static class MsOfficeExportRepoFactoryProvider
    {
        public static MsOfficeExportRepoFactory__Interface? Current { get; set; }

        /// <summary>Returns the configured factory, or throws a clear error if a host forgot to set it.</summary>
        public static MsOfficeExportRepoFactory__Interface Require() =>
            Current ?? throw new System.InvalidOperationException(
                "No MsOfficeExportRepoFactory has been configured. The host application must set " +
                "MsOfficeExportRepoFactoryProvider.Current at startup (see DL.MsOfficeApi_Impl).");
    }
}
