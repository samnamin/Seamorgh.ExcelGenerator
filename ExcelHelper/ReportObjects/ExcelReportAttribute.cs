using System;

namespace ExcelHelper.ReportObjects
{
    public class ExcelReportAttribute : Attribute
    {
        public ExcelReportAttribute(bool visable = true, string name="", bool useInFormula=true)
        {
            Visible = visable;
            Name = name;
            UseInFormulas = useInFormula;
        }

        public bool Visible { get; set; }
        public string Name { get; set; }
        public bool UseInFormulas  { get; set; }
        public Type Type { get; set; }
    }
}
