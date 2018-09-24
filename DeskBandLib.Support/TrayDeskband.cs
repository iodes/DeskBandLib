using DeskBandLib.Support.Interop;
using System;
using System.Runtime.InteropServices;

namespace DeskBandLib.Support
{
    class TrayDeskband : IDisposable
    {
        public ITrayDeskband Instance { get; private set; }

        private static Type TrayDeskbandType { get; set; }

        public TrayDeskband()
        {
            if (TrayDeskbandType == null)
                TrayDeskbandType = Type.GetTypeFromCLSID(new Guid("E6442437-6C68-4f52-94DD-2CFED267EFB9"));

            Instance = Activator.CreateInstance(TrayDeskbandType) as ITrayDeskband;
        }

        public void Dispose()
        {
            if (Instance != null && Marshal.IsComObject(Instance))
            {
                Marshal.ReleaseComObject(Instance);
            }
        }
    }
}
