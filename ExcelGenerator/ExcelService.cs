using ClosedXML.Excel;
using ExcelHelper.Reports.ExcelReports;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelGenerator
{
    public static class ExcelService
    {
        public static ExcelGeneratedFileResult GenerateExcel(WorkBook workBook)
        {
            var fakeReport = new List<string> { "Ahmad", "Zabih", "Ramin", "Marzieh", "Reza", "Ahmad" };

            try
            {
                //-------------------------------------------
                //  Create Workbook (integrated with using statement)
                //-------------------------------------------
                using var wb = new XLWorkbook();

                // Create Sheet
                var firstSheet = wb.Worksheets.Add("First");

                // Fill Headers (for First sheet)
                firstSheet.Cell(1, 1).Value = "Id";
                firstSheet.Cell(1, 2).Value = "FirstName";
                firstSheet.Cell(1, 3).Value = "LastName";

                // Fill Data (for First sheet)
                for (int index = 1; index <= fakeReport.Count; index++)
                {
                    firstSheet.Cell(index + 1, 1).Value = fakeReport[index - 1];
                    firstSheet.Cell(index + 1, 2).Value = fakeReport[index - 1];
                    firstSheet.Cell(index + 1, 3).Value = fakeReport[index - 1];
                }

                // Save
                using var stream = new MemoryStream();
                wb.SaveAs(stream);
                var content = stream.ToArray();
                return new ExcelGeneratedFileResult { Content = content, FileName = workBook.FileName };
            }
            catch (Exception e)
            {
                // ignored
                throw;
            }
        }
    }
}