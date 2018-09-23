using System;

namespace DeskBandLib.Attributes
{
    /// <summary>
    /// The display name of the extension and the description for the HelpText(displayed in status bar when menu command selected).
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class DeskBandInfoAttribute : Attribute
    {
        public string DisplayName { get; }

        public string HelpText { get; }

        public DeskBandInfoAttribute()
        {

        }

        public DeskBandInfoAttribute(string displayName, string helpText)
        {
            DisplayName = displayName;
            HelpText = helpText;
        }
    }
}
