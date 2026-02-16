using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Group4333.Models;

namespace Group4333.Excel
{
    public class ExcelImporter
    {
        /// <summary>
        /// Импорт данных из Excel файла 1.xlsx
        /// </summary>
        public List<Service> ImportFromFile(string filePath)
        {
            List<Service> services = new List<Service>();

            // Проверяем, существует ли файл
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Файл Excel не найден", filePath);
            }

            // Устанавливаем лицензию для EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Берем первый лист
                var worksheet = package.Workbook.Worksheets[0];

                // Определяем количество строк с данными
                int rowCount = worksheet.Dimension.Rows;

                // Читаем данные со 2-й строки (1-я обычно заголовки)
                for (int row = 2; row <= rowCount; row++)
                {
                    Service service = new Service
                    {
                        Id = Convert.ToInt32(worksheet.Cells[row, 1].Value ?? 0),
                        Name = worksheet.Cells[row, 2].Value?.ToString() ?? "",
                        Type = worksheet.Cells[row, 3].Value?.ToString() ?? "",
                        Price = Convert.ToDecimal(worksheet.Cells[row, 4].Value ?? 0)
                    };

                    services.Add(service);
                }
            }

            return services;
        }

        /// <summary>
        /// Группировка услуг по виду
        /// </summary>
        public Dictionary<string, List<Service>> GroupByType(List<Service> services)
        {
            var grouped = new Dictionary<string, List<Service>>();

            foreach (var service in services)
            {
                if (!grouped.ContainsKey(service.Type))
                {
                    grouped[service.Type] = new List<Service>();
                }
                grouped[service.Type].Add(service);
            }

            // Сортируем каждую группу по возрастанию цены
            foreach (var type in grouped.Keys)
            {
                grouped[type] = grouped[type].OrderBy(s => s.Price).ToList();
            }

            return grouped;
        }
    }
}