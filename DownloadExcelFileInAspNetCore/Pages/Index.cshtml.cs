using ClosedXML.Excel;
using DownloadExcelFileInAspNetCore.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DownloadExcelFileInAspNetCore.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
        }

        public ActionResult OnGetDownloadExcel()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee() { EmployeeId = 1, EmployeeName = "Mohammed Gadi", Age = 26, Gender = "Male", Designation = "Sofware Engineer" });
            employees.Add(new Employee() { EmployeeId = 2, EmployeeName = "Zareen Khan", Age = 30, Gender = "Female", Designation = "CTO" });
            employees.Add(new Employee() { EmployeeId = 3, EmployeeName = "John Doe", Age = 26, Gender = "Male", Designation = "Senior Sofware Engineer" });

            var titles = new List<string[]>() { new string[] { "EmployeeId", "EmployeeName", "Age", "Gender", "Designation" } };
            var byteArray = ExportToExcel(employees, "Employees", titles);
            var fileName = "Employees.xlsx";
            return new JsonResult(new { data = Convert.ToBase64String(byteArray), fileName });
        }

        public byte[] ExportToExcel(List<Employee> list, string worksheetTitle, List<string[]> titles)
        {
            var wb = new XLWorkbook(); //create workbook
            var ws = wb.Worksheets.Add(worksheetTitle); //add worksheet to workbook

            var rangeTitle = ws.Cell(1, 1).InsertData(titles); //insert titles to first row
            rangeTitle.AddToNamed("Titles");
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true; //font must be bold

            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle; //attach style to the range

            if (list != null && list.Count() > 0)
            {
                //insert data to from second row on
                ws.Cell(2, 1).InsertData(list);
                ws.Columns().AdjustToContents();
            }

            //save file to memory stream and return it as byte array
            using (var memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}