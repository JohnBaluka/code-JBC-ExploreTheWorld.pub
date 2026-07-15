using System;
using System.Runtime.InteropServices;

namespace JBC.ExploreTheWorld.AL.WinFormsLib
{
    /// <summary>
    /// Shared utility for Watcher forms — obtains a running COM application object.
    /// Mirrors VBA <c>GetObject(, "Word.Application")</c>.
    /// </summary>
    internal static class WatcherComHelper
    {
#if !NETFRAMEWORK
        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
#endif

        /// <summary>
        /// Returns the running COM object for <paramref name="progId"/>.
        /// Throws <see cref="InvalidOperationException"/> if the application is not running.
        /// </summary>
        internal static object GetActiveCom(string progId)
        {
#if NETFRAMEWORK
            return Marshal.GetActiveObject(progId);
#else
            var type = Type.GetTypeFromProgID(progId)
                ?? throw new InvalidOperationException(
                    $"ProgID '{progId}' not found. Is the application installed?");
            var clsid = type.GUID;
            int hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            if (hr != 0)
                Marshal.ThrowExceptionForHR(hr);
            return obj;
#endif
        }

        /// <summary>
        /// Returns the running COM object for <paramref name="progId"/>, or launches a new
        /// instance if the application is not already running. <paramref name="created"/> is
        /// true when a new instance was started (the caller should make it visible).
        /// </summary>
        internal static object GetOrCreateCom(string progId, out bool created)
        {
            try
            {
                created = false;
                return GetActiveCom(progId);
            }
            catch
            {
                var type = Type.GetTypeFromProgID(progId)
                    ?? throw new InvalidOperationException(
                        $"ProgID '{progId}' not found. Is the application installed?");
                created = true;
                return Activator.CreateInstance(type)!;
            }
        }
    }
}
