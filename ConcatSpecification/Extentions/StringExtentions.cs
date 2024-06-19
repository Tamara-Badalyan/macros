using System.Text.RegularExpressions;

namespace ConcatSpecification.Extentions;

public static class StringExtentions
{
    public static string Join(this string firstString, string secondString)
    {
        return string.Join(", ", new[] { firstString, secondString }.Where(s => !string.IsNullOrEmpty(s)));
    }

    public static bool CompareWithoutSpaces(this string firstString, string secondString)
    {
        string pattern = @"^\d+(\.\d+)*\s*";
        firstString = Regex.Replace(firstString, pattern, "");
        secondString = Regex.Replace(secondString, pattern, "");

        firstString = Regex.Replace(firstString, @" ", "").Trim();
        secondString = Regex.Replace(secondString, @" ", "").Trim();

        return string.Equals(firstString, secondString, StringComparison.OrdinalIgnoreCase);
    }
}
