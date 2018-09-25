using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Interop
{
    /// <summary>
    /// Exposes a method that is used to communicate focus changes for a user input object contained in the Shell.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F1DB8392-7331-11D0-8C99-00A0C92DBFE8")]
    public interface IInputObjectSite
    {
        /// <summary>
        /// Informs the browser that the focus has changed.
        /// </summary>
        /// <param name="punkObj">The address of the IUnknown interface of the object gaining or losing the focus.</param>
        /// <param name="fSetFocus">Indicates if the object has gained or lost the focus. If this value is nonzero, the object has gained the focus.
        /// If this value is zero, the object has lost the focus.</param>
        /// <returns>Returns S_OK if the method was successful, or a COM-defined error code otherwise.</returns>
        [PreserveSig]
        int OnFocusChangeIS([MarshalAs(UnmanagedType.IUnknown)] object punkObj, int fSetFocus);
    }
}
