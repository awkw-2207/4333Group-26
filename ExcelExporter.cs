using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Group4333.Models;

namespace Group4333.Excel
{
    public class ExcelExporter
    {
        /// <summary>
        /// Экспорт данных в Excel с разделением по листам
        /// </summary>
        public void ExportToFile(Dictionary<string, List<Service>> groupedServices, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                foreach (var group in groupedServices)
                {
                    // Создаем лист для каждого вида услуги
                    // Обрезаем название листа до 31 символа (ограничение Excel)
                    string sheetName = group.Key.Length > 31 ? group.Key.Substring(0, 31) : group.Key;
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // Заголовки
                    worksheet.Cells[1, 1].Value = "Id";
                    worksheet.Cells[1, 2].Value = "Название услуги";
                    worksheet.Cells[1, 3].Value = "Стоимость";

                    // Делаем заголовки жирными
                    using (var range = worksheet.Cells[1, 1, 1, 3])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // Данные
                    int row = 2;
                    foreach (var service in group.Value)
                    {
                        worksheet.Cells[row, 1].Value = service.Id;
                        worksheet.Cells[row, 2].Value = service.Name;
                        worksheet.Cells[row, 3].Value = service.Price;

                        // Формат для цены
                        worksheet.Cells[row, 3].Style.Numberformat.Format = "#,##0.00 ₽";

                        row++;
                    }

                    // Автоподбор ширины колонок
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }

                // Сохраняем файл
                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}