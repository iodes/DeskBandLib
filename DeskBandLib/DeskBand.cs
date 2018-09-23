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

        #region Variables
        protected IInputObjectSite DeskBandSite;
        #endregion

        #region Properties
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

        #region IObjectWithSite
        public void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite)
        {
            if (DeskBandSite != null)
                Marshal.ReleaseComObject(DeskBandSite);

            DeskBandSite = (IInputObjectSite)pUnkSite;
        }

        public void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite)
        {
            ppvSite = DeskBandSite;
        }
        #endregion

        #region IDeskBand2
        public virtual int CanRenderComposited(out bool pfCanRenderComposited)
        {
            pfCanRenderComposited = true;
            return S_OK;
        }

        public int SetCompositionState(bool fCompositionEnabled)
        {
            fCompositionEnabled = true;
            return S_OK;
        }

        public int GetCompositionState(out bool pfCompositionEnabled)
        {
            pfCompositionEnabled = false;
            return S_OK;
        }

        public int GetBandInfo(uint dwBandID, DESKBANDINFO.DBIF dwViewMode, ref DESKBANDINFO pdbi)
        {
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

        public int GetWindow(out IntPtr phwnd)
        {
            phwnd = Handle;
            return S_OK;
        }

        public int ContextSensitiveHelp(bool fEnterMode)
        {
            return S_OK;
        }

        public int ShowDW([In] bool fShow)
        {
            if (fShow)
                Show();
            else
                Hide();

            return S_OK;
        }

        public int CloseDW([In] uint dwReserved)
        {
            Dispose(true);
            return S_OK;
        }

        public int ResizeBorderDW(RECT prcBorder, [In, MarshalAs(UnmanagedType.IUnknown)] IntPtr punkToolbarSite, bool fReserved)
        {
            return E_NOTIMPL;
        }
        #endregion

        #region Register / Unregister
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            string guid = t.GUID.ToString("B");

            var deskBandInfo = t.GetCustomAttributes(typeof(DeskBandInfoAttribute), false) as DeskBandInfoAttribute[];

            // Register only the extension if the attribute DeskBandInfo is used.
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

                // TaskBar
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
    }
}
