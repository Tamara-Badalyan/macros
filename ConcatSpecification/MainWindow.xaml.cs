using ConcatSpecification;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
            var excelFilePaths = new List<string>();

            var excelOpenFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выбор excel файлов",
                Filter = "Excel файл|*.xlsx"
            };


            if (excelOpenFileDialog.ShowDialog().HasValue)
                excelFilePaths = excelOpenFileDialog.FileNames.ToList();

            ExcelFileBuilder.Build(excelFilePaths);

            //FilterByNotepad.ReadTXT(filterPath);
            //FilterByNotepad.FilterExcel(excelPath);

            this.Close();
        }
    }
}
