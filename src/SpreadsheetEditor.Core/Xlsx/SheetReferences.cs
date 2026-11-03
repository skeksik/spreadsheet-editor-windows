using System.Text;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Finds the sheet names a formula refers to: <c>Name!A1</c>, <c>'Two words'!A1</c> (with <c>''</c> for a quote), the 3-D forms
/// <c>First:Last!A1</c> and <c>'First:Last'!A1</c>, and external forms such as <c>[1]Name!A1</c> (the name after the brackets). It is a
/// conservative scanner, not a formula parser: the writer uses it to refuse a change that would leave references to a renamed or deleted
/// sheet (docs/spec.md §9.30), so it errs towards finding a reference.
/// </summary>
internal static class SheetReferences
{
    private const string Separators = " \t\r\n+-*/^&=<>(),;{}%!\"";

    /// <summary>The sheet names referred to, in order, possibly repeated.</summary>
    /// <param name="text">A formula, or any text that may hold formulas (a part's XML).</param>
    /// <param name="skipStrings">True for a formula: text in double quotes is a string, not a reference. False for raw XML.</param>
    public static List<string> Find(string text, bool skipStrings)
    {
        var names = new List<string>();
        var i = 0;
        while (i < text.Length)
        {
            var c = text[i];
            if (skipStrings && c == '"')
            {
                i = text.IndexOf('"', i + 1);
                i = i < 0 ? text.Length : i + 1;
                continue;
            }

            if (c == '\'')
            {
                var end = QuotedEnd(text, i, out var quoted);
                if (end < text.Length && text[end] == '!')
                {
                    AddParts(names, quoted);
                }

                i = end;
                continue;
            }

            if (Separators.Contains(c, StringComparison.Ordinal))
            {
                i++;
                continue;
            }

            var start = i;
            while (i < text.Length && !Separators.Contains(text[i], StringComparison.Ordinal) && text[i] != '\'')
            {
                i++;
            }

            if (i < text.Length && text[i] == '!')
            {
                var token = text[start..i];
                var bracket = token.LastIndexOf(']');
                AddParts(names, bracket >= 0 ? token[(bracket + 1)..] : token);
            }
        }

        return names;
    }

    /// <summary>True when the text refers to the sheet, compared without case as sheet names are.</summary>
    public static bool Mentions(string text, string sheetName, bool skipStrings) =>
        Find(text, skipStrings).Exists(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// The formula with every reference to sheet <paramref name="from"/> made a reference to <paramref name="to"/>, quoted. Used for the
    /// defined names copied with a duplicated sheet. A 3-D reference is left as it is.
    /// </summary>
    public static string Replace(string formula, string from, string to)
    {
        var result = new StringBuilder(formula.Length);
        var replacement = "'" + to.Replace("'", "''", StringComparison.Ordinal) + "'";
        var i = 0;
        while (i < formula.Length)
        {
            var c = formula[i];
            if (c == '"')
            {
                var close = formula.IndexOf('"', i + 1);
                close = close < 0 ? formula.Length : close + 1;
                result.Append(formula, i, close - i);
                i = close;
                continue;
            }

            if (c == '\'')
            {
                var end = QuotedEnd(formula, i, out var quoted);
                var match = end < formula.Length && formula[end] == '!' && string.Equals(quoted, from, StringComparison.OrdinalIgnoreCase);
                result.Append(match ? replacement : formula[i..end]);
                i = end;
                continue;
            }

            if (Separators.Contains(c, StringComparison.Ordinal))
            {
                result.Append(c);
                i++;
                continue;
            }

            var start = i;
            while (i < formula.Length && !Separators.Contains(formula[i], StringComparison.Ordinal) && formula[i] != '\'')
            {
                i++;
            }

            var token = formula[start..i];
            var isSheet = i < formula.Length && formula[i] == '!' && string.Equals(token, from, StringComparison.OrdinalIgnoreCase);
            result.Append(isSheet ? replacement : token);
        }

        return result.ToString();
    }

    /// <summary>The index after a quoted name that starts at <paramref name="start"/>, and the name without quotes.</summary>
    private static int QuotedEnd(string text, int start, out string name)
    {
        var builder = new StringBuilder();
        var i = start + 1;
        while (i < text.Length)
        {
            if (text[i] == '\'')
            {
                if (i + 1 < text.Length && text[i + 1] == '\'')
                {
                    builder.Append('\'');
                    i += 2;
                    continue;
                }

                i++;
                break;
            }

            builder.Append(text[i]);
            i++;
        }

        name = builder.ToString();
        return i;
    }

    private static void AddParts(List<string> names, string reference)
    {
        var bracket = reference.LastIndexOf(']');
        if (bracket >= 0)
        {
            reference = reference[(bracket + 1)..];
        }

        foreach (var part in reference.Split(':'))
        {
            if (part.Length > 0)
            {
                names.Add(part);
            }
        }
    }
}
