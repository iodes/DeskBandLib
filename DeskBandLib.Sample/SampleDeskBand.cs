using DeskBandLib.Attributes;
using System;
using System.Drawing;
using System.Runtime.InteropServices;

/// <summary>
/// To use, you must sign with the same key as the DeskBandLib project.
/// To display the background transparently, you must use a Black background.
/// </summary>

namespace DeskBandLib.Sample
{
    [ComVisible(true)]
    [Guid("0FC2816C-6A60-485E-A886-DDC9CA7ADD61")]
    [DeskBandInfo("Sample Desk Band", "DeskBandLib Sample Application")]
    public partial class SampleDeskBand : DeskBand
    {
        public SampleDeskBand()
        {
            Title = "Sample Desk Band";
            MinSize = new Size(200, 40);
            MinSizeVertical = new Size(100, 40);
            InitializeComponent();
        }
    }
}
