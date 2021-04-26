using ExcelHelper.Reports.ExcelReports;
using System.ComponentModel;
using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public abstract class PropertyOption
    {
        public PropertyOption(Location startLocation)
        {
            StartLocation = startLocation;
        }

        public Location StartLocation { get; set; }
        public Color BackColor { get; set; } = Color.White;
        public Color ForeColor { get; set; } = Color.Black;
        [DefaultValue(true)]
        public bool Visible { get; set; }
    }
}
