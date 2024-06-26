using ConcatSpecification.Extentions;
using ConcatSpecification.Models;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ConcatSpecification;

public class ExcelFileBuilder
{
    private static readonly string folderName = "Templates";

    public static void Build(List<string> filePaths, string selectedFile)
    {      
        var productSheetModels = ConvertToModels(filePaths);

        var resultModel = ConcatSpecification(productSheetModels, selectedFile);

        BuildExcel(resultModel);
    }

    /// <summary>
    /// Проходит через лист Excel, преобразуя каждую строку в соответствующую модель.
    /// </summary>
    private static List<ProductSheetModel> ConvertToModels(List<string> filePaths)
    {
        var currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var excelApp = new Application();

        var productSheets = new List<ProductSheetModel>();

        foreach (var filePath in filePaths)
        {
            var fullFilePath = Path.Combine(currentDirectory, filePath);

            var workbook = excelApp.Workbooks.Open(fullFilePath);
            var worksheet = (Worksheet)workbook.Sheets[1];
            //#2 проверяет таблицу: 9шт. столбцов
            if (worksheet.UsedRange.Columns.Count != 9) throw new Exception("Column count must be 9");

            var usedRange = worksheet.UsedRange;

            //#2 проверяет таблицу: угол таблицы с ячейки А1
            if (usedRange.Cells[1, 1] == null) throw new Exception("Cell A1 must not be empty");

            var productSheet = ProductSheetModel.Create(Path.GetFileName(fullFilePath), worksheet.Name);
            ProductGroupModel productGroupModel = default;

            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                var groupName = usedRange.Cells[row, 2].GetValue();

                if (string.IsNullOrEmpty(groupName))
                {
                    continue;
                }

                //проверяет начало новой группы
                if (ProductGroupModel.IsGroupName(groupName) && string.IsNullOrEmpty(usedRange.Cells[row, 7].GetValue()))
                {
                    productGroupModel = ProductGroupModel.Create(groupName);
                    productSheet.ProductGroupModels.Add(productGroupModel);
                    continue;
                }
                
                var col = 1;
                var product = new ProductModel
                {
                    Position = usedRange.Cells[row, col++].GetValue(),
                    Title = usedRange.Cells[row, col++].GetValue(),
                    Type = usedRange.Cells[row, col++].GetValue(),
                    Code = usedRange.Cells[row, col++].GetValue(),
                    Provider = usedRange.Cells[row, col++].GetValue(),
                    Measurement = usedRange.Cells[row, col++].GetValue(),
                    Count = usedRange.Cells[row, col++].GetValue(),
                    Weight = usedRange.Cells[row, col++].GetValue(),
                    Note = usedRange.Cells[row, col++].GetValue()
                };
                productGroupModel.Products.Add(product);
              
            }
            productSheets.Add(productSheet);

            workbook.Close();
            excelApp.Quit();
        }

        return productSheets;
    }

    private static ProductSheetModel ConcatSpecification(List<ProductSheetModel> productSheetModels, string selectedFile)
    {
        var firstSheetModel = productSheetModels.First(i => i.FilePath == selectedFile);
        productSheetModels.Remove(firstSheetModel);

        var resultModel = ProductSheetModel.Create("result", firstSheetModel.SheetName);

        foreach (var productSheetModel in productSheetModels)
        {

            foreach (var productGroupModel in productSheetModel.ProductGroupModels)
            {
                var group = firstSheetModel.ProductGroupModels.FirstOrDefault(i => i.CompareGroupName(productGroupModel));
                var groupIndex = firstSheetModel.ProductGroupModels.IndexOf(group);

                if (group is null) continue;

                foreach (var product in productGroupModel.Products)
                {
                    //#6 сравнение и поиск по молеляи
                    var matchingProduct = group.Products.FirstOrDefault(p => p.Equals(product));

                    if (matchingProduct is null)
                    {
                        product.IsAdded = true;
                        group.Products.Add(product);
                        continue;
                    }

                    var mainCount = string.IsNullOrEmpty(matchingProduct.Count) ? 0 : double.Parse(matchingProduct.Count.Replace(',', '.'));
                    var count = string.IsNullOrEmpty(matchingProduct.Count) ? 0 : double.Parse(product.Count.Replace(',', '.'));

                    //#7 если для комплекта разные значения (в столбец G (“Кол.”))
                    //и входят в комплект электрощитового оборудования (с тире), добавляется  новая позиция 
                    if (!matchingProduct.IsSerialized && mainCount != count)
                    {
                        group.Products.Add(product);
                    }

                    //#7 при одинаковости строк добавляется значение в столбец G(“Кол.”).
                    var total = mainCount + count;
                    matchingProduct.Count = total.ToString().Replace('.', ',');

                    //#6 при одинаковости строк (с исключением столбцов A,G,H,I):
                    //добавляется текст в столбцы A, I, H, при условии если текст разный в этих ячейках.
                    if (!String.Equals(matchingProduct.Position, product.Position, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingProduct.Position = matchingProduct.Position.Join(product.Position);
                    }

                    if (!String.Equals(matchingProduct.Note, product.Note, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingProduct.Note = matchingProduct.Note.Join(product.Note);
                    }

                    if (!String.Equals(matchingProduct.Weight, product.Weight, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingProduct.Weight = matchingProduct.Weight.Join(product.Weight);
                    }

                }

                firstSheetModel.ProductGroupModels[groupIndex] = group;
            }
        }

        return firstSheetModel;
    }


    private static void BuildExcel(ProductSheetModel productSheetModels)
    {
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

        var templatePath = Path.Combine(currentDirectory, folderName, "Template");

        string newFileName = "результат.xlsx";
        string newFilePath = Path.Combine(currentDirectory, newFileName);

        var excelApp = new Application();

        var workbook = excelApp.Workbooks.Open(templatePath);
        var worksheet = (Worksheet)workbook.ActiveSheet;

        var row = 2;
        var rowHeghtWithPoints = 0.45 * 72;

        //#10 Изменение ширины столбцов
        ((Range)worksheet.Columns["A"]).ColumnWidth= 9.29;
        ((Range)worksheet.Columns["B"]).ColumnWidth= 65;
        ((Range)worksheet.Columns["C"]).ColumnWidth= 29.57;
        ((Range)worksheet.Columns["D"]).ColumnWidth= 17;
        ((Range)worksheet.Columns["E"]).ColumnWidth= 22;
        ((Range)worksheet.Columns["F"]).ColumnWidth= 9.29;
        ((Range)worksheet.Columns["G"]).ColumnWidth= 9.29;
        ((Range)worksheet.Columns["H"]).ColumnWidth= 11.86;
        ((Range)worksheet.Columns["I"]).ColumnWidth= 19.57;
        

        foreach (var group in productSheetModels.ProductGroupModels)
        {
            worksheet.Cells[row, 2] = group.GroupName;

            Range cellForGroup = (Range)worksheet.Cells[row++, 2];
            cellForGroup.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cellForGroup.VerticalAlignment = XlVAlign.xlVAlignCenter;
            cellForGroup.Font.Bold = true;

            foreach (var product in group.Products)
            {
                Range cell = (Range)worksheet.Rows[row];
                cell.RowHeight = rowHeghtWithPoints;
                cell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                cell.Font.Bold = false;


                if (product.IsAdded)
                {
                    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }

                int col = 1;
                worksheet.Cells[row, col++] = product.Position;
                worksheet.Cells[row, col++] = product.Title;
                worksheet.Cells[row, col++] = product.Type;
                worksheet.Cells[row, col++] = product.Code;
                worksheet.Cells[row, col++] = product.Provider;
                worksheet.Cells[row, col++] = product.Measurement;
                worksheet.Cells[row, col++] = product.Count;
                worksheet.Cells[row, col++] = product.Weight;
                worksheet.Cells[row, col++] = product.Note;

                row++;
            }
            row += 2;
        }

        if (File.Exists(newFilePath))
        {
            File.Delete(newFilePath);
        }

        workbook.SaveAs(newFilePath);
        workbook.Close();
        excelApp.Quit();

        ProcessStartInfo psInfo = new ProcessStartInfo
        {
            FileName = newFilePath,
            UseShellExecute = true
        };
        Process.Start(psInfo);
    }

}
