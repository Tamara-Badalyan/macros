namespace ConcatSpecification.Models;

/// <summary>
/// Represent a single excel sheet
/// </summary>
public class ProductSheetModel
{
    /// <summary>
    /// Excel file name
    /// </summary>
    public string SheetName { get; set; }
    /// <summary>
    /// Grouped products list
    /// </summary>
    public List<ProductGroupModel> ProductGroupModels {get;set;}

    public static ProductSheetModel Create(string source)
    {
        return new ProductSheetModel { SheetName = source, ProductGroupModels = new List<ProductGroupModel>() };
    }
}
