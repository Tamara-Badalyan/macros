namespace ConcatSpecification.Models;

/// <summary>
/// Represent a single excel sheet
/// </summary>
public class ProductSheetModel
{

    /// <summary>
    /// Excel file Path
    /// </summary>
    public string FilePath { get; set; }

    /// <summary>
    /// Excel sheet name
    /// </summary>
    public string SheetName { get; set; }
    /// <summary>
    /// Grouped products list
    /// </summary>
    public List<ProductGroupModel> ProductGroupModels {get;set;}

    public static ProductSheetModel Create(string filePath, string source)
    {
        return new ProductSheetModel { FilePath = filePath, SheetName = source, ProductGroupModels = new List<ProductGroupModel>() };
    }
}
