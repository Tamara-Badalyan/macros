using System.Text.RegularExpressions;
using System.Windows.Shapes;

namespace ConcatSpecification.Models;

public class ProductGroupModel
{
    public const string patternForGroupName = @"\d\.(?!\d)[\p{L}\s]+[\p{L}\s]*";

    public string GroupName { get; set; }
    public List<ProductModel> Products { get; set; }

    public static bool IsGroupName(string name)
    {
        var regex = new Regex(patternForGroupName);
        return regex.IsMatch(name);
    }

    public bool CompareGroupName(ProductGroupModel productGroup)
    {
        string pattern = @"^\d+\.\s*";
        var currentName = Regex.Replace(this.GroupName, pattern, "");
        var productGroupName = Regex.Replace(productGroup.GroupName, pattern, "");

        currentName = Regex.Replace(currentName, @"\s+", " ").Trim();
        productGroupName = Regex.Replace(productGroupName, @"\s+", " ").Trim();

        return string.Equals(currentName, productGroupName, StringComparison.OrdinalIgnoreCase);
    }

    public static ProductGroupModel Create(string groupName)
    {
        return new ProductGroupModel { GroupName = groupName, Products = new List<ProductModel>() };
    }
}
