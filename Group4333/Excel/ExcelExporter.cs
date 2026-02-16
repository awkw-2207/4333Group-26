using Group4333.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace Group4333.Excel
{
    public class ExcelExporter
    {
        public void ExportToFile(Dictionary<string, List<Service>> groupedServices, string filePath)
        {

            using (var package = new ExcelPackage())
            {
                foreach (var group in groupedServices)
                {
                    string sheetName = group.Key.Length > 31 ? group.Key.Substring(0, 31) : group.Key;
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    worksheet.Cells[1, 1].Value = "Id";
                    worksheet.Cells[1, 2].Value = "Название услуги";
                    worksheet.Cells[1, 3].Value = "Стоимость";

                    using (var range = worksheet.Cells[1, 1, 1, 3])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    int row = 2;
                    foreach (var service in group.Value)
                    {
                        worksheet.Cells[row, 1].Value = service.Id;
                        worksheet.Cells[row, 2].Value = service.Name;
                        worksheet.Cells[row, 3].Value = service.Price;

                        worksheet.Cells[row, 3].Style.Numberformat.Format = "#,##0.00 ₽";

                        row++;
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}