using System.Collections.Generic;

namespace ExcelHelper.Reports.ExcelReports
{
    public class Sheet
    {
        public Sheet(string caption, string name)
        {
            Name = name;
            Caption = caption;
            Rows = new List<Row>();
            Columns = new List<Column>();
            Tables = new List<Table>();
            MergedCells = new List<string>();
        }

        public string Name { get; set; }
        public string Caption { get; set; }
        public List<Row> Rows { get; set; }
        public List<Column> Columns { get; set; }
        public List<Table> Tables { get; set; }
        public List<string> MergedCells { get; set; }
    }
}
