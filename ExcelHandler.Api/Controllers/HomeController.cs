using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDataReader;
using ExcelHandler.Api.Database;
using ExcelHandler.Api.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExcelHandler.Api.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private AppDbContext _context;

        public HomeController(AppDbContext context)
        {
            _context = context;
        }

        [HttpPost]
        public  async Task<IActionResult> UploadFile(IFormFile file)
        {
            var fileName = file.FileName;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
            {

                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };  

                var dataSet = reader.AsDataSet(conf);

                var dataTable = dataSet.Tables[0];


                for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                    Employee employee =new Employee
                    {
                        FirstName = dataTable.Rows[i][0].ToString(),
                        LastName = dataTable.Rows[i][1].ToString(),
                        Salary = Decimal.Parse(dataTable.Rows[i][2].ToString())
                    };

                    _context.employees.Add(employee);
                }

            }

            await _context.SaveChangesAsync();
            var result = Getemployees().Result.Value;
            return Ok(result);
        }

        [HttpGet]
        public async Task<IActionResult> ExportExcel()
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                // Set author for excel file
                package.Workbook.Properties.Author = "Aspodel";
                // Set title for excel file
                package.Workbook.Properties.Title = "Employee List";
                // Add comment to excel file
                package.Workbook.Properties.Comments = "Hello (^_^)";
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = "No";
                worksheet.Cells[1, 2].Value = "First Name";
                worksheet.Cells[1, 3].Value = "Last Name";
                worksheet.Cells[1, 4].Value = "Salary";

                // Style for Excel 
                worksheet.DefaultColWidth = 16;
                worksheet.Cells.Style.Font.Size = 16;


                //Export Data from Table employees
                var list = await _context.employees.ToListAsync();
                for (int i = 0; i < list.Count; i++)
                {
                    var item = list[i];
                    worksheet.Cells[i + 2, 1].Value = i + 1;
                    worksheet.Cells[i + 2, 2].Value = item.FirstName;
                    worksheet.Cells[i + 2, 3].Value = item.LastName;
                    worksheet.Cells[i + 2, 4].Value = item.Salary;
                }

                fileContents = package.GetAsByteArray();
            }

            if(fileContents == null || fileContents.Length == 0)
            {
                return NoContent();
            }

            return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "Employees.xlsx");
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Employee>>> Getemployees()
        {
            return await _context.employees.ToListAsync();
        }


        [HttpDelete]
        public async Task<ActionResult<Employee>> DeleteAll()
        {
            var all = from allEmployees in _context.employees select allEmployees;
            _context.employees.RemoveRange(all);
            await _context.SaveChangesAsync();

            var result = Getemployees().Result.Value;
            return Ok(result);
        }
    }
}
