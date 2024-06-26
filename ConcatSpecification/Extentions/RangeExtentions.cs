using Range = Microsoft.Office.Interop.Excel.Range;

namespace ConcatSpecification.Extentions;

public static class RangeExtentions
{
    public static string GetValue(this object range)
    {
        if (range is Range)
        {
            return Convert.ToString((range as Range).Value2);
        }
        throw new InvalidOperationException();
    }
}
