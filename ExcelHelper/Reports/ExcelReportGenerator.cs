using ExcelHelper.Reports.ExcelReports;
using ExcelHelper.Reports.ExcelReports.Template;
using ExcelHelper.VoucherStatementReport;

namespace ExcelHelper.Reports
{
    public class ExcelReportGenerator
    {
        public void VoucherStatementExcelReport(VoucherStatementPageResult result)
        {

            var file = new WorkBook("FileName", "Path");
            var sheet1 = SheetTemplates.VoucherStatementTemplate(result);
            file.Sheets.Add(sheet1);
        }
    }
}
