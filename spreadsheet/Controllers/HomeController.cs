using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
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
        public ActionResult UpdateExcel(List<List<string>> updatedExcelData, List<List<string>> newData, List<string> newRow, int? clearRowIndex, bool? addRow)
        {
            try
            {
                if (updatedExcelData != null)
                {
                    // Clear existing data
                    ViewBag.ExcelData = new List<List<string>>();

                    // Add new data if provided
                    if (newData != null && newData.Count > 0)
                    {
                        ViewBag.ExcelData.AddRange(newData);
                    }

                    // Update existing data
                    if (updatedExcelData != null && updatedExcelData.Count > 0)
                    {
                        ViewBag.ExcelData.AddRange(updatedExcelData);
                    }

                    // Add new row if requested
                    if (addRow.HasValue && addRow.Value)
                    {
                        ViewBag.ExcelData.Add(newRow);
                    }

                    // Clear row if requested
                    if (clearRowIndex.HasValue && clearRowIndex.Value >= 0 && clearRowIndex.Value < ViewBag.ExcelData.Count)
                    {
                        ClearRow(ViewBag.ExcelData, clearRowIndex.Value);
                    }
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



        private void ClearRow(List<List<string>> data, int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < data.Count)
            {
                for (int i = 0; i < data[rowIndex].Count; i++)
                {
                    data[rowIndex][i] = string.Empty;
                }
            }
        }

        private void AddRow(List<List<string>> data, List<string> newRow)
        {
            data.Add(newRow);
        }

    }
}