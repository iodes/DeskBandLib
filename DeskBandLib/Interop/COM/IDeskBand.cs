using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Interop
{
    /// <summary>
    /// Gets information about a band object.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("EB0FE172-1A3A-11D0-89B3-00A0C90A90AC")]
    interface IDeskBand : IDockingWindow
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

        /// <summary>
        /// Gets state information for a band object.
        /// </summary>
        /// <param name="dwBandID">The identifier of the band, assigned by the container. The band object can retain this value if it is required.</param>
        /// <param name="dwViewMode">The view mode of the band object. One of the following values: DBIF_VIEWMODE_NORMAL, DBIF_VIEWMODE_VERTICAL, DBIF_VIEWMODE_FLOATING, DBIF_VIEWMODE_TRANSPARENT.</param>
        /// <param name="pdbi">Pointer to a DESKBANDINFO structure that receives the band information for the object. The dwMask member of this structure indicates the specific information that is being requested.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int GetBandInfo(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi);
    }
}
