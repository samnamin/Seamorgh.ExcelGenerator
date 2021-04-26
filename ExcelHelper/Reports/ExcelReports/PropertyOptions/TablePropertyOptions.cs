namespace ExcelHelper.Reports.ExcelReports.PropertyOptions
{
    public class TablePropertyOptions : PropertyOption
    {
        public TablePropertyOptions(Location startLocation) : base(startLocation) { }
        public bool IsBordered { get; set; }
        public int BorderSize { get; set; }
        public string BorderType { get; set; }
    }
}
