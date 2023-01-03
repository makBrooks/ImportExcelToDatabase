using ImportExcelToDatabase.Models;
using ImportExcelToDatabase.Repository;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcelToDatabase.Controllers
{
    public class HomeController : Controller
    {
        private readonly IExceltoDatabaseRepository _excelRepo;
        public HomeController(IExceltoDatabaseRepository excelRepo)
        {
            _excelRepo = excelRepo;
        }

        public async Task<IActionResult> Index()
        {
            List<ExceltoDatabaseDetails> obj = await _excelRepo.Get();
            return View(obj);
        }
        [HttpPost]
        public IActionResult Index(IFormFile File)
        {
            try
            {
                if (File != null && File.Length > 0)
                {
                    string fileExtension = File.ContentType.Trim();
                    if ((fileExtension == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") || (fileExtension == "application/vnd.ms-excel"))
                    {

                        var filename = File.FileName + DateTime.Now.ToString("yymmssfff");
                        var path = Path.Combine(
                                    Directory.GetCurrentDirectory(), "wwwroot/Files",
                                    filename);
                        using (var stream = new FileStream(path, FileMode.Create))
                        {
                            File.CopyToAsync(stream);
                        }

                        var fileinfo = new FileInfo(path);
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        using (var package = new ExcelPackage(fileinfo))
                        {
                            var workbook = package.Workbook;
                            var worksheet = workbook.Worksheets.First();
                            int colCount = worksheet.Dimension.End.Column;  //get Column Count
                            int rowCount = worksheet.Dimension.End.Row;
                            List<ExceltoDatabase> excelsheet = new List<ExceltoDatabase>();


                            for (int row = 2; row <= rowCount; row++)
                            {
                                ExceltoDatabase excel = new ExceltoDatabase();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Value?.ToString().Trim()))
                                    {
                                        if (col == 1)
                                        {
                                            excel.slno = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 2)
                                        {
                                            excel.name = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 3)
                                        {
                                            excel.email = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 4)
                                        {
                                            excel.mobile = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                        else if (col == 5)
                                        {
                                            excel.specialization = worksheet.Cells[row, col].Value?.ToString().Trim();
                                        }
                                    }
                                }
                                if (!string.IsNullOrWhiteSpace(excel.slno))
                                {
                                    excelsheet.Add(excel);
                                }
                            }
                            if (excelsheet != null)
                                _excelRepo.AddExcelDetails(excelsheet);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
