using DeskBandLib.Attributes;
using DeskBandLib.Interop;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DeskBandLib
{
    public class DeskBand : UserControl, IObjectWithSite, IDeskBand2
    {
        #region Constants
        private const int S_OK = 0;
        private const int E_NOTIMPL = unchecked((int)0x80004001);
        #endregion

        #region Properties
        /// <summary>
        /// Gets the ID of the current DeskBand.
        /// </summary>
        protected uint DeskBandID { get; private set; }

        /// <summary>
        /// Gets the Site of the current DeskBand.
        /// </summary>
        protected IInputObjectSite DeskBandSite { get; private set; }

        /// <summary>
        /// Title of the band object, displayed by default on the left or top of the object.
        /// </summary>
        [Browsable(true)]
        [DefaultValue("")]
        public string Title { get; set; }

        /// <summary>
        /// Determines if the title appears in the DeskBand.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(false)]
        public bool ShowTitle { get; set; }

        /// <summary>
        /// Determines if the deskband has a fixed position and size, and if the gripper is shown.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(false)]
        public bool IsFixed { get; set; }

        /// <summary>
        /// Determines if the height of the horizontal deskband is allowed to change. For a deskband in the vertical orientation, it will be the width.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(true)]
        public bool HeightCanChange { get; set; } = true;

        /// <summary>
        /// Minimum size of the band object. Default value of -1 sets no minimum constraint.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(typeof(Size), "-1,-1")]
        public Size MinSize { get; set; }

        /// <summary>
        /// Maximum size of the band object. Default value of -1 sets no maximum constraint.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(typeof(Size), "-1,-1")]
        public Size MaxSize { get; set; }

        /// <summary>
        /// Minimum vertical size of the band object. Default value of -1 sets no maximum constraint. (Used when the taskbar is aligned vertically.)
        /// </summary>
        [Browsable(true)]
        [DefaultValue(typeof(Size), "-1,-1")]
        public Size MinSizeVertical { get; set; }

        /// <summary>
        /// Says that band object's size must be multiple of this size. Defauilt value of -1 does not set this constraint.
        /// </summary>
        [Browsable(true)]
        [DefaultValue(typeof(Size), "-1,-1")]
        public Size IntegralSize { get; set; }
        #endregion

        #region Initialize
        public DeskBand()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Name = "DeskBand";
        }
        #endregion

        #region Implements

        #region Methods
        private int GetWindowImpl(out IntPtr phwnd)
        {
            phwnd = Handle;
            return S_OK;
        }

        private int ContextSensitiveHelpImpl(bool fEnterMode)
        {
            return S_OK;
        }

        private int ShowDWImpl(bool fShow)
        {
            if (fShow)
            {
                Show();
            }
            else
            {
                Hide();
            }

            return S_OK;
        }

        private int CloseDWImpl(uint dwReserved)
        {
            Dispose(true);
            return S_OK;
        }

        private int ResizeBorderDWImpl(RECT prcBorder, IntPtr punkToolbarSite, bool fReserved)
        {
            return E_NOTIMPL;
        }

        private int GetBandInfoImpl(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi)
        {
            DeskBandID = dwBandID;

            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_MINSIZE))
            {
                if (dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_FLOATING) || dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_VERTICAL))
                {
                    pdbi.ptMinSize.Y = MinSizeVertical.Width;
                    pdbi.ptMinSize.X = MinSizeVertical.Height;
                }
                else
                {
                    pdbi.ptMinSize.X = MinSize.Width;
                    pdbi.ptMinSize.Y = MinSize.Height;
                }
            }
            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_MAXSIZE))
            {
                if (dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_FLOATING) || dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_VERTICAL))
                {
                    pdbi.ptMaxSize.Y = MaxSize.Width;
                    pdbi.ptMaxSize.X = MaxSize.Height;
                }
                else
                {
                    pdbi.ptMaxSize.X = MaxSize.Width;
                    pdbi.ptMaxSize.Y = MaxSize.Height;
                }
            }
            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_INTEGRAL))
            {
                if (dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_FLOATING) || dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_VERTICAL))
                {
                    pdbi.ptIntegral.Y = IntegralSize.Width;
                    pdbi.ptIntegral.X = IntegralSize.Height;
                }
                else
                {
                    pdbi.ptIntegral.X = IntegralSize.Width;
                    pdbi.ptIntegral.Y = IntegralSize.Height;
                }
            }

            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_ACTUAL))
            {
                if (dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_FLOATING) || dwViewMode.HasFlag(DESKBANDINFO.DBIF.DBIF_VIEWMODE_VERTICAL))
                {
                    pdbi.ptActual.Y = Size.Width;
                    pdbi.ptActual.X = Size.Height;
                }
                else
                {
                    pdbi.ptActual.X = Size.Width;
                    pdbi.ptActual.Y = Size.Height;
                }
            }

            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_TITLE))
            {
                pdbi.wszTitle = Title;
                pdbi.dwMask &= ShowTitle ? DESKBANDINFO.DBIM.DBIM_TITLE : ~DESKBANDINFO.DBIM.DBIM_TITLE;
            }

            if (pdbi.dwMask.HasFlag(DESKBANDINFO.DBIM.DBIM_MODEFLAGS))
            {
                pdbi.dwModeFlags = DESKBANDINFO.DBIMF.DBIMF_NORMAL;
                pdbi.dwModeFlags |= IsFixed ? DESKBANDINFO.DBIMF.DBIMF_FIXED | DESKBANDINFO.DBIMF.DBIMF_NOGRIPPER : 0;
                pdbi.dwModeFlags |= HeightCanChange ? DESKBANDINFO.DBIMF.DBIMF_VARIABLEHEIGHT : 0;
                pdbi.dwModeFlags &= ~DESKBANDINFO.DBIMF.DBIMF_BKCOLOR;
            }

            return S_OK;
        }

        private int CanRenderCompositedImpl(out bool pfCanRenderComposited)
        {
            pfCanRenderComposited = true;
            return S_OK;
        }

        private int SetCompositionStateImpl(bool fCompositionEnabled)
        {
            return S_OK;
        }

        private int GetCompositionState(out bool pfCompositionEnabled)
        {
            pfCompositionEnabled = false;
            return S_OK;
        }
        #endregion

        #region IObjectWithSite
        void IObjectWithSite.SetSite(object pUnkSite)
        {
            if (DeskBandSite != null)
                Marshal.ReleaseComObject(DeskBandSite);

            DeskBandSite = (IInputObjectSite)pUnkSite;
        }

        void IObjectWithSite.GetSite(ref Guid riid, out object ppvSite)
        {
            ppvSite = DeskBandSite;
        }
        #endregion

        #region IOleWindow
        int IOleWindow.GetWindow(out IntPtr phwnd)
        {
            return GetWindowImpl(out phwnd);
        }

        int IOleWindow.ContextSensitiveHelp(bool fEnterMode)
        {
            return ContextSensitiveHelpImpl(fEnterMode);
        }
        #endregion

        #region IDockingWindow
        int IDockingWindow.GetWindow(out IntPtr phwnd)
        {
            return GetWindowImpl(out phwnd);
        }

        int IDockingWindow.ContextSensitiveHelp(bool fEnterMode)
        {
            return ContextSensitiveHelpImpl(fEnterMode);
        }

        int IDockingWindow.ShowDW(bool fShow)
        {
            return ShowDWImpl(fShow);
        }

        int IDockingWindow.CloseDW(uint dwReserved)
        {
            return CloseDWImpl(dwReserved);
        }

        int IDockingWindow.ResizeBorderDW(RECT prcBorder, IntPtr punkToolbarSite, bool fReserved)
        {
            return ResizeBorderDWImpl(prcBorder, punkToolbarSite, fReserved);
        }
        #endregion

        #region IDeskBand
        int IDeskBand.GetWindow(out IntPtr phwnd)
        {
            return GetWindowImpl(out phwnd);
        }

        int IDeskBand.ContextSensitiveHelp(bool fEnterMode)
        {
            return ContextSensitiveHelpImpl(fEnterMode);
        }

        int IDeskBand.ShowDW(bool fShow)
        {
            return ShowDWImpl(fShow);
        }

        int IDeskBand.CloseDW(uint dwReserved)
        {
            return CloseDWImpl(dwReserved);
        }

        int IDeskBand.ResizeBorderDW(RECT prcBorder, IntPtr punkToolbarSite, bool fReserved)
        {
            return ResizeBorderDWImpl(prcBorder, punkToolbarSite, fReserved);
        }

        int IDeskBand.GetBandInfo(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi)
        {
            return GetBandInfoImpl(dwBandID, dwViewMode, ref pdbi);
        }
        #endregion

        #region IDeskBand2
        int IDeskBand2.GetWindow(out IntPtr phwnd)
        {
            return GetWindowImpl(out phwnd);
        }

        int IDeskBand2.ContextSensitiveHelp(bool fEnterMode)
        {
            return ContextSensitiveHelpImpl(fEnterMode);
        }

        int IDeskBand2.ShowDW(bool fShow)
        {
            return ShowDWImpl(fShow);
        }

        int IDeskBand2.CloseDW(uint dwReserved)
        {
            return CloseDWImpl(dwReserved);
        }

        int IDeskBand2.ResizeBorderDW(RECT prcBorder, IntPtr punkToolbarSite, bool fReserved)
        {
            return ResizeBorderDWImpl(prcBorder, punkToolbarSite, fReserved);
        }

        int IDeskBand2.GetBandInfo(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi)
        {
            return GetBandInfoImpl(dwBandID, dwViewMode, ref pdbi);
        }

        int IDeskBand2.CanRenderComposited(out bool pfCanRenderComposited)
        {
            return CanRenderCompositedImpl(out pfCanRenderComposited);
        }

        int IDeskBand2.SetCompositionState(bool fCompositionEnabled)
        {
            return SetCompositionStateImpl(fCompositionEnabled);
        }

        int IDeskBand2.GetCompositionState(out bool pfCompositionEnabled)
        {
            return GetCompositionState(out pfCompositionEnabled);
        }
        #endregion

        #endregion

        #region COM Methods
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            string guid = t.GUID.ToString("B");

            var deskBandInfo = t.GetCustomAttributes(typeof(DeskBandInfoAttribute), false) as DeskBandInfoAttribute[];

            if (deskBandInfo.Length == 1)
            {
                var rkClass = Registry.ClassesRoot.CreateSubKey(@"CLSID\" + guid);
                var rkCat = rkClass.CreateSubKey("Implemented Categories");

                string _displayName = t.Name;
                string _helpText = t.Name;


                if (deskBandInfo[0].DisplayName != null)
                {
                    _displayName = deskBandInfo[0].DisplayName;
                }

                if (deskBandInfo[0].HelpText != null)
                {
                    _helpText = deskBandInfo[0].HelpText;
                }

                rkClass.SetValue(null, _displayName);
                rkClass.SetValue("MenuText", _displayName);
                rkClass.SetValue("HelpText", _helpText);

                rkCat.CreateSubKey("{00021492-0000-0000-C000-000000000046}");

                Console.WriteLine(string.Format("{0} {1} {2}", guid, _displayName, "successfully registered."));
            }
            else
            {
                Console.WriteLine(guid + " has no attributes");
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            string guid = t.GUID.ToString("B");

            var deskBandInfo = t.GetCustomAttributes(typeof(DeskBandInfoAttribute), false) as DeskBandInfoAttribute[];

            if (deskBandInfo.Length == 1)
            {
                string _displayName = t.Name;

                if (deskBandInfo[0].DisplayName != null)
                {
                    _displayName = deskBandInfo[0].DisplayName;
                }

                Registry.ClassesRoot.CreateSubKey(@"CLSID").DeleteSubKeyTree(guid);

                Console.WriteLine(string.Format("{0} {1} {2}", guid, _displayName, "successfully removed."));
            }
            else
            {
                Console.WriteLine(guid + " has no attributes");
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Hide the current DeskBand from the taskbar.
        /// </summary>
        public void HideDeskBand()
        {
            if (DeskBandSite != null)
            {
                var bandSite = DeskBandSite as IBandSite;
                bandSite.RemoveBand(DeskBandID);
            }
        }
        #endregion
    }
}
