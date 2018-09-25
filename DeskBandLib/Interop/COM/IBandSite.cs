using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Interop
{
    /// <summary>
    /// Exposes methods that control band objects.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CF504B0-DE96-11D0-8B3F-00A0C911E8E5")]
    public interface IBandSite
    {
        /// <summary>
        /// Adds a band to a band site object.
        /// </summary>
        /// <param name="punk">The interface of a band site object.</param>
        /// <returns>Returns the band ID in ShortFromResult(hresult).</returns>
        [PreserveSig]
        int AddBand(ref object punk);

        /// <summary>
        /// Enumerates the bands in a band site.
        /// </summary>
        /// <param name="uBand">Call the method with this parameter starting at 0 to begin enumerating. If this parameter is -1, the pdwBandIDparameter is ignored and this method returns the count of the bands in the band site.</param>
        /// <param name="pdwBandID">The address of a band ID variable that receives the band ID.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code for errors. If the first parameter is -1, the count of the bands in the band site is returned.</returns>
        [PreserveSig]
        int EnumBands(int uBand, out uint pdwBandID);

        /// <summary>
        /// Gets information about a band in a band site.
        /// </summary>
        /// <param name="dwBandID">The ID of the band object to query.</param>
        /// <param name="ppstb">Address of an IDeskBand interface pointer that, when this method returns successfully, points to the IDeskBand object that represents the band. This value can be NULL.</param>
        /// <param name="pdwState">Pointer to a DWORD value that, when this method returns successfully, receives the state of the band object. This state is a combination of BSSF_VISIBLE, BSSF_NOTITLE, and BSSF_UNDELETEABLE. See BANDSITEINFO for more information on those flags. This value can be NULL if the state information is not needed.</param>
        /// <param name="pszName">Pointer to a buffer of cchName Unicode characters that, when this method returns successfully, receives the name of the band object.</param>
        /// <param name="cchName">The size of the pszName buffer, in characters.</param>
        /// <returns>If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.</returns>
        [PreserveSig]
        int QueryBand(uint dwBandID, out IDeskBand ppstb, out BANDSITEINFO.BSSF pdwState, [MarshalAs(UnmanagedType.LPWStr)] out string pszName, int cchName);

        /// <summary>
        /// Set the state of a band in the band site.
        /// </summary>
        /// <param name="dwBandID">The ID of the band to set. If this parameter is -1, then set the state of all bands in the band site.</param>
        /// <param name="dwMask">The mask of the states to set.</param>
        /// <param name="dwState">The state values to be set. These are combinations of BSSF_* flags. For more information, see BANDSITEINFO.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int SetBandState(uint dwBandID, BANDSITEINFO.BSIM dwMask, BANDSITEINFO.BSSF dwState);

        /// <summary>
        /// Removes a band from the band site.
        /// </summary>
        /// <param name="dwBandID">The ID of the band to remove.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int RemoveBand(uint dwBandID);

        /// <summary>
        /// Gets a specified band object from a band site.
        /// </summary>
        /// <param name="dwBandID">The ID of the band object to get.</param>
        /// <param name="riid">The IID of the object to obtain.</param>
        /// <param name="ppv">The address of a pointer variable that receives a pointer to the object requested.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int GetBandObject(uint dwBandID, ref Guid riid, out IntPtr ppv);

        /// <summary>
        /// Sets information about the band site.
        /// </summary>
        /// <param name="pbsinfo">The address of a BANDSITEINFO structure that receives the band site information for the object. The dwMask member of this structure specifies what information is being set.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int SetBandSiteInfo([In] ref BANDSITEINFO pbsinfo);

        /// <summary>
        /// Gets information about a band in the band site.
        /// </summary>
        /// <param name="pbsinfo">The address of a BANDSITEINFO structure that contains the band site information for the object. The dwMask member of this structure specifies what information is being requested.</param>
        /// <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int GetBandSiteInfo([In, Out] ref BANDSITEINFO pbsinfo);
    }
}
