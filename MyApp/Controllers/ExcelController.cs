using ExcelHelper.ReportObjects;
using ExcelHelper.VoucherStatementReport;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace MyApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("ExportExcel")]
        public IActionResult ExportExcel()
        {
            // TODO: Dear Shahab, please replace my naive example with something real and more practical to build an Excel like you sent to me
            var arg = new VoucherStatementPageResult
            {
                ReportName = "TestReport",
                SummaryAccounts = new List<SummaryAccount>
                {
                    new SummaryAccount {AccountName = "MyAccount"},
                    new SummaryAccount {AccountName = "MyAccount2"}
                },
                RowResult = new List<VoucherStatementRowResult>
                {
                    new VoucherStatementRowResult
                    {
                        AccountCode = "Code1",
                        Credit = 2342,
                        Debit = 232
                    },
                    new VoucherStatementRowResult
                    {
                        AccountCode = "Code2",
                        Credit = 222,
                        Debit = 23333
                    }
                }
            };

            var result = ExcelReportGenerator.VoucherStatementExcelReport(arg);

            return File(result.Content, result.ContentType, result.FileName);
        }
    }
}