using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Support.Interop
{
    /// <summary>
    /// Exposes methods that show, hide, and query deskbands.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6D67E846-5B9C-4db8-9CBC-DDE12F4254F1")]
    public interface ITrayDeskband
    {
        /// <summary>
        /// Shows a specified deskband.
        /// </summary>
        /// <param name="clsid">A reference to a deskband CLSID.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int ShowDeskBand([In, MarshalAs(UnmanagedType.Struct)] ref Guid clsid);

        /// <summary>
        /// Hides a specified deskband.
        /// </summary>
        /// <param name="clsid">A reference to a deskband CLSID.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int HideDeskBand([In, MarshalAs(UnmanagedType.Struct)] ref Guid clsid);

        /// <summary>
        /// Indicates whether a deskband is shown.
        /// </summary>
        /// <param name="clsid">A reference to a deskband CLSID.</param>
        /// <returns>Returns S_OK if the deskband is shown, S_FALSE if the deskband is not shown, or an error value otherwise.</returns>
        [PreserveSig]
        int IsDeskBandShown([In, MarshalAs(UnmanagedType.Struct)] ref Guid clsid);

        /// <summary>
        /// Refreshes the deskband registration cache.
        /// </summary>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int DeskBandRegistrationChanged();
    }
}
