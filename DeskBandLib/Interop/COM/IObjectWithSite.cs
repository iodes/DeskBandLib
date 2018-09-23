using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Interop
{
    /// <summary>
    /// Provides a simple way to support communication between an object and its site in the container.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")]
    public interface IObjectWithSite
    {
        /// <summary>
        /// Enables a container to pass an object a pointer to the interface for its site.
        /// </summary>
        /// <param name="pUnkSite">A pointer to the IUnknown interface pointer of the site managing this object.
        /// If NULL, the object should call Release on any existing site at which point the object no longer knows its site.</param>
        void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);

        /// <summary>
        /// Retrieves the latest site passed using SetSite.
        /// </summary>
        /// <param name="riid">The IID of the interface pointer that should be returned in ppvSite.</param>
        /// <param name="ppvSite">Address of pointer variable that receives the interface pointer requested in riid.</param>
        void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite);
    }
}
