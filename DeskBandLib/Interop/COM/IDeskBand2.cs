using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Interop
{
    /// <summary>
    /// Exposes methods to enable and query translucency effects in a deskband object.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("79D16DE4-ABEE-4021-8D9D-9169B261D657")]
    public interface IDeskBand2 : IDeskBand
    {
        #region IOleWindow
        [PreserveSig]
        new int GetWindow(out IntPtr phwnd);

        [PreserveSig]
        new int ContextSensitiveHelp(bool fEnterMode);
        #endregion

        #region IDockingWindow
        [PreserveSig]
        new int ShowDW([In] bool fShow);

        [PreserveSig]
        new int CloseDW([In] uint dwReserved);

        [PreserveSig]
        new int ResizeBorderDW(RECT prcBorder, [In, MarshalAs(UnmanagedType.IUnknown)] IntPtr punkToolbarSite, bool fReserved);
        #endregion

        #region IDeskBand
        [PreserveSig]
        new int GetBandInfo(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi);
        #endregion

        /// <summary>
        /// Indicates the deskband's ability to be displayed as translucent.
        /// </summary>
        /// <param name="pfCanRenderComposited">When this method returns, contains a BOOL indicating ability.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int CanRenderComposited(out bool pfCanRenderComposited);

        /// <summary>
        /// Sets the composition state.
        /// </summary>
        /// <param name="fCompositionEnabled">TRUE to enable the composition state; otherwise, FALSE.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int SetCompositionState(bool fCompositionEnabled);

        /// <summary>
        /// Gets the composition state.
        /// </summary>
        /// <param name="pfCompositionEnabled">When this method returns, contains a BOOL that indicates state.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int GetCompositionState(out bool pfCompositionEnabled);
    }
}
