using System;
using System.Data;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Column
    {
        public Column(Location location)
        {
            Location = location;
        }
        public DataColumn Data { get; set; }
        public string Name { get; set; }
        public Type Type { get; set; }
        public object Value { get; set; }
        public Location Location { get; set; }
        public bool Wordwrap { get; set; }
        public TextAlign Align { get; set; } = TextAlign.rtl;
        public int Width { get; set; } = 20;
        public Category Category { get; set; } = Category.General;
        public bool Visible { get; set; }
        public bool AutoFill { get; set; }
    }

    public enum TextAlign
    {
        rtl,
        ltr,
        center
    }

    public enum Category
    {
        General,
        Number,
        Currency,
        Date,
        Time,
        Percentage,
        Text,
        Custom
    }
}
