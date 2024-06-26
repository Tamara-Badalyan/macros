using ConcatSpecification;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ExcelFileSelector;

public partial class FileSelectionWindow : Window
{
    public List<string> SelectedFiles { get; set; }

    public FileSelectionWindow(List<string> selectedFiles)
    {
        InitializeComponent();
        SelectedFiles = selectedFiles;
        FilesComboBox.ItemsSource = SelectedFiles.Select(i => System.IO.Path.GetFileName(i)).ToList(); ;
        if (SelectedFiles.Count > 0)
        {
            FilesComboBox.SelectedIndex = 0;
        }
    }

    private void ConfirmSelection_Click(object sender, RoutedEventArgs e)
    {
        if (FilesComboBox.SelectedItem != null)
        {
            string selectedFileName = FilesComboBox.SelectedItem.ToString();
            ExcelFileBuilder.Build(SelectedFiles, selectedFileName);
        }
        else
        {
            MessageBox.Show("Please select a file from the dropdown.");
        }
    }
}