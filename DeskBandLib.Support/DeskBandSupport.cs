using System;

namespace DeskBandLib.Support
{
    public static class DeskBandSupport
    {
        private const int S_OK = 0;
        private const int S_FALSE = 1;

        public static bool ShowDeskBand(Guid clsid)
        {
            using (var trayDeskBand = new TrayDeskband())
            {
                if (trayDeskBand.Instance.IsDeskBandShown(ref clsid) == S_FALSE)
                {
                    bool result = false;
                    trayDeskBand.Instance.DeskBandRegistrationChanged();
                    result = trayDeskBand.Instance.ShowDeskBand(ref clsid) == S_OK;
                    trayDeskBand.Instance.DeskBandRegistrationChanged();

                    return result;
                }
            }

            return false;
        }

        public static bool HideDeskBand(Guid clsid)
        {
            using (var trayDeskBand = new TrayDeskband())
            {
                if (trayDeskBand.Instance.IsDeskBandShown(ref clsid) == S_OK)
                {
                    bool result = false;
                    trayDeskBand.Instance.DeskBandRegistrationChanged();
                    result = trayDeskBand.Instance.HideDeskBand(ref clsid) == S_OK;
                    trayDeskBand.Instance.DeskBandRegistrationChanged();

                    return result;
                }
            }

            return false;
        }

        public static bool IsDeskBandShown(Guid clsid)
        {
            using (var trayDeskBand = new TrayDeskband())
            {
                if (trayDeskBand.Instance.IsDeskBandShown(ref clsid) == S_OK)
                {
                    return true;
                }

                return false;
            }
        }

        public static bool DeskBandRegistrationChanged()
        {
            using (var trayDeskBand = new TrayDeskband())
            {
                if (trayDeskBand.Instance.DeskBandRegistrationChanged() == S_OK)
                {
                    return true;
                }

                return false;
            }
        }
    }
}
