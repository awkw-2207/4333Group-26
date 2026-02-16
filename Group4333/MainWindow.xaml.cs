using Group4333.Database;
using Group4333.Excel;
using Group4333.Models;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Linq;
using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        private DatabaseHelper dbHelper;
        private ExcelImporter excelImporter;
        private ExcelExporter excelExporter;
        private List<Service> currentServices;
        public MainWindow()
        {
            InitializeComponent();
            dbHelper = new DatabaseHelper();
            excelImporter = new ExcelImporter();
            excelExporter = new ExcelExporter();

            try
            {
                dbHelper.CreateTable();
                UpdateStatus("Готов к работе", "Gray");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Ошибка инициализации БД: {ex.Message}", "Red");
            }


        }
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.Title = "Выберите файл 1.xlsx";

                if (openFileDialog.ShowDialog() == true)
                {
                    UpdateStatus("Импортируем данные...", "Blue");

                    currentServices = excelImporter.ImportFromFile(openFileDialog.FileName);

                    dbHelper.SaveServices(currentServices);

                    ServicesListBox.ItemsSource = currentServices;

                    UpdateStatus($"Импортировано {currentServices.Count} услуг", "Green");

                    var grouped = excelImporter.GroupByType(currentServices);
                    string stats = string.Join(", ", grouped.Keys.Select(k => $"{k}: {grouped[k].Count}"));
                    MessageBox.Show($"Импорт завершен!\n\nГруппы:\n{stats}", "Успех");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Ошибка: {ex.Message}", "Red");
                MessageBox.Show($"Ошибка при импорте: {ex.Message}", "Ошибка");
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var services = dbHelper.GetAllServices();

                if (services.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте данные!", "Предупреждение");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "export_services.xlsx";
                saveFileDialog.Title = "Сохранить файл Excel";

                if (saveFileDialog.ShowDialog() == true)
                {
                    UpdateStatus("Экспортируем данные...", "Blue");

                    var grouped = excelImporter.GroupByType(services);

                    excelExporter.ExportToFile(grouped, saveFileDialog.FileName);

                    UpdateStatus($"Экспорт завершен: {saveFileDialog.FileName}", "Green");

                    MessageBox.Show($"Экспорт завершен!\n\nФайл сохранен:\n{saveFileDialog.FileName}", "Успех");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Ошибка: {ex.Message}", "Red");
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка");
            }
        }

        private void UpdateStatus(string message, string color)
        {
            StatusText.Text = message;
            StatusText.Foreground = color switch
            {
                "Red" => System.Windows.Media.Brushes.Red,
                "Green" => System.Windows.Media.Brushes.Green,
                "Blue" => System.Windows.Media.Brushes.Blue,
                _ => System.Windows.Media.Brushes.Gray
            };
        }
    }
}