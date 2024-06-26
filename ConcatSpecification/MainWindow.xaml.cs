using ExcelFileSelector;
using Microsoft.Win32;
using System.Windows;

namespace ConcatSpecification
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnConcatSpecification_Click(object sender, RoutedEventArgs e)
        {
            var excelFileNames = new List<string>();

            var excelOpenFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выбор excel файлов",
                Filter = "Excel файл|*.xlsx"
            };


            if (excelOpenFileDialog.ShowDialog().HasValue)
            {
                excelFileNames = excelOpenFileDialog.FileNames.Select(i => System.IO.Path.GetFileName(i)).ToList(); ;
                FileSelectionWindow fileSelectionWindow = new FileSelectionWindow(excelOpenFileDialog.FileNames.ToList());
                fileSelectionWindow.ShowDialog();
            }


            this.Close();
        }
    }
}
