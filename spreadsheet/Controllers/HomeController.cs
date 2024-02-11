using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using spreadsheet.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using OfficeOpenXml;

namespace spreadsheet.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(IFormFile file)
        {
            try
            {
                if (file != null && file.Length > 0)
                {
                    string fileName = Path.GetFileName(file.FileName);
                    string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadedFiles", fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    List<List<string>> excelData = ReadExcel(filePath);

                    ViewBag.ExcelData = excelData;
                }
                else
                {
                    ViewBag.ErrorMessage = "Please upload a file.";
                }
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "Error: " + ex.Message;
            }

            return View("Index");
        }

        private List<List<string>> ReadExcel(string filePath)
        {
            List<List<string>> excelData = new List<List<string>>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    List<string> rowData = new List<string>();
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Text;
                        rowData.Add(cellValue);
                    }

                    excelData.Add(rowData);
                }
            }

            return excelData;
        }

        [HttpPost]
        public ActionResult UpdateExcel(List<List<string>> updatedExcelData)
        {
            try
            {
                if (updatedExcelData != null)
                {
                    ViewBag.ExcelData = updatedExcelData;
                }
                else
                {
                    ViewBag.ErrorMessage = "No data received for update.";
                }
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "Error: " + ex.Message;
            }

            return View("Index");
        }
    }
}
