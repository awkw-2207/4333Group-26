using Group4333.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace Group4333.Excel
{
    public class ExcelImporter
    {
        public List<Service> ImportFromFile(string filePath)
        {
            List<Service> services = new List<Service>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Файл Excel не найден", filePath);
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

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

            foreach (var type in grouped.Keys)
            {
                grouped[type] = grouped[type].OrderBy(s => s.Price).ToList();
            }

            return grouped;
        }
    }
}