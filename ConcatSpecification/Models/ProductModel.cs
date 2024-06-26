using ConcatSpecification.Extentions;

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

    //#6 сравнение с исключением столбцов A,G,H,I
    public bool Equals(ProductModel otherItem)
    {
        return Type.CompareWithoutSpaces(otherItem.Type)
         && Code.CompareWithoutSpaces(otherItem.Code)
         && Provider.CompareWithoutSpaces(otherItem.Provider)
         && Measurement.CompareWithoutSpaces(otherItem.Measurement)
         && Title.CompareWithoutSpaces(otherItem.Title);
    }

}