using Microsoft.Office.Interop.Excel;
using System.IO;
using System;
using Range = Microsoft.Office.Interop.Excel.Range;


namespace ConcatSpecification
{
    public class FilterByNotepad
    {
        public static List<string> keywords = new List<string>();

        public static void ReadTXT(string filterPath)
        {
            keywords = new List<string>();

            string[] lines = File.ReadAllLines(filterPath);
            foreach (string line in lines)
                keywords.Add(line.ToLower());
        }

        public static void FilterExcel(string excelPath)
        {
            List<List<string>> result = new List<List<string>>();
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = excelPath;
            string filePath = Path.Combine(currentDirectory, fileName);

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath);
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            //int rowCount = range.Rows.Count;
            //int colCount = range.Columns.Count;

            foreach (var keyword in keywords)
            {
                for (int i = 1; i <= 2; i++)
                {
                    string cellValue = Convert.ToString((range.Cells[i, 5] as Range).Value2);
                    if (!string.IsNullOrEmpty(cellValue) && cellValue.ToLower().Contains(keyword))
                    {
                        if (result.Any(x => x.Contains(Convert.ToString((range.Cells[i, 1] as Range).Value2))))
                            continue;

                        result.Add(new List<string> {
                            Convert.ToString((range.Cells[i, 1] as Range).Value2),
                            Convert.ToString((range.Cells[i, 5] as Range).Value2),
                            Convert.ToString((range.Cells[i, 11] as Range).Value2)
                        });
                    }
                }
            }

            //workbook.Close();
            //excelApp.Quit();

            SaveToExcel(result);
        }

        public static void SaveToExcel(List<List<string>> data)
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];

            Range columnRange = (Range)worksheet.Columns[2];
            columnRange.ColumnWidth = 100;
            columnRange.WrapText = true;

            for (int i = 0; i < data.Count; i++)
            {
                List<string> rowData = data[i];
                for (int j = 0; j < rowData.Count; j++)
                {
                    worksheet.Cells[i + 1, j + 1] = rowData[j];
                }
            }

            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string newFileName = "результат.xlsx";
            string newFilePath = Path.Combine(currentDirectory, newFileName);
            workbook.SaveAs(newFilePath);
            workbook.Close();
            excelApp.Quit();

            System.Diagnostics.Process.Start(newFilePath);
        }
    }
}
