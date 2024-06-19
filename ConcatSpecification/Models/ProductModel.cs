using ConcatSpecification.Extentions;
using System.Text.RegularExpressions;

namespace ConcatSpecification.Models;

public class ProductModel : IEquatable<ProductModel>
{
    public string Position { get; set; }//A
    public string Title { get; set; }//B
    public string Type { get; set; }//C
    public string Code { get; set; }//D
    public string Provider { get; set; }//E
    public string Measurement { get; set; }//F
    public string Count { get; set; }//G
    public string Weight { get; set; }//H
    public string Note { get; set; }//I
    public bool IsAdded { get; set; }
    public bool IsSerialized => !Title.Trim().StartsWith('-');

    public override bool Equals(object other)
    {
        var otherItem = other as ProductModel;

        if (otherItem == null)
            return false;

        return Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Title, Count);
    }

    public bool Equals(ProductModel otherItem)
    {
        if(Code == "80251")
        {
            var test1 = Type.CompareWithoutSpaces(otherItem.Type);
            var test2 = Code.CompareWithoutSpaces(otherItem.Code);
            var test3 = Provider.CompareWithoutSpaces(otherItem.Provider);
            var test4 = Measurement.CompareWithoutSpaces(otherItem.Measurement);
            var test5 = Title.CompareWithoutSpaces(otherItem.Title);
        }
        

        return Type.CompareWithoutSpaces(otherItem.Type)
         && Code.CompareWithoutSpaces(otherItem.Code)
         && Provider.CompareWithoutSpaces(otherItem.Provider)
         && Measurement.CompareWithoutSpaces(otherItem.Measurement)
         && Title.CompareWithoutSpaces(otherItem.Title);
    }

}
//    static bool CompareTitles(string line1, string line2)
//    {
//        string pattern = @"^\d+(\.\d+)*\s*";
//        line1 = Regex.Replace(line1, pattern, "");
//        line2 = Regex.Replace(line2, pattern, "");

//        line1 = Regex.Replace(line1, @" ", "").Trim();
//        line2 = Regex.Replace(line2, @" ", "").Trim();

//        return string.Equals(line1, line2, StringComparison.OrdinalIgnoreCase);
//    }
//}
