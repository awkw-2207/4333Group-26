using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void AuthorInfoButton_Click(object sender, RoutedEventArgs e)
        {
            Window window = new Window();
            window.Show();
            this.Close();
        }
    }
}