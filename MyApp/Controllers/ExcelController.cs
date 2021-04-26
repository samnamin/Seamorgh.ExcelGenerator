using ExcelGenerator;
using ExcelHelper.Reports.ExcelReports;
using Microsoft.AspNetCore.Mvc;

namespace MyApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("ExportExcel")]
        public IActionResult ExportExcel()
        {
            var wb = new WorkBook("name", "");

            var result = ExcelService.GenerateExcel(wb);

            return File(result.Content, result.ContentType, result.FileName);
        }
    }
}