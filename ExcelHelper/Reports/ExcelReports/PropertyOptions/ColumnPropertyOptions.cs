using System.Drawing;

namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class ColumnPropertyOptions : PropertyOption
    {
        public ColumnPropertyOptions(Location startLocation) : base(startLocation) { }
        public string EndLocation { get; set; }
        public Size Size { get; set; }
    }
}
