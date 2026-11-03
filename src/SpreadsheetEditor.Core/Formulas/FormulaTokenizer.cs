namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// Splits A1-style formula text into tokens, enough to find the references in it. It does not check that the formula is valid and
/// never evaluates anything: text it does not recognise becomes <see cref="FormulaTokenKind.Other"/>, and the tokens together cover the
/// whole text, so copying every token's text gives the formula back unchanged. A leading <c>=</c> is allowed and is an Other token.
/// </summary>
internal static class FormulaTokenizer
{
    /// <summary>Whether <paramref name="c"/> may appear in a name, an unquoted sheet name or a function name.</summary>
    public static bool IsNameChar(char c) => char.IsLetterOrDigit(c) || c is '_' or '.' or '\\' or '?';

    public static List<FormulaToken> Tokenize(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        var tokens = new List<FormulaToken>();
        var i = 0;
        var lastBracketEnd = -1;
        while (i < formula.Length)
        {
            var c = formula[i];
            int end;
            if (c == '"')
            {
                end = SkipQuoted(formula, i, '"');
                tokens.Add(new FormulaToken(FormulaTokenKind.String, i, end - i));
            }
            else if (c == '[')
            {
                end = SkipBracketed(formula, i);
                tokens.Add(new FormulaToken(FormulaTokenKind.Bracketed, i, end - i));
                lastBracketEnd = end;
            }
            else if (c == '\'')
            {
                end = SkipQuoted(formula, i, '\'');
                if (end < formula.Length && formula[end] == '!' && end - i >= 2)
                {
                    var name = formula[(i + 1)..(end - 1)].Replace("''", "'", StringComparison.Ordinal);
                    // A sheet name cannot hold '[' (SheetNames), so a quoted prefix with one names another workbook.
                    var external = name.Contains('[', StringComparison.Ordinal) || i == lastBracketEnd;
                    var multi = external || name.Contains(':', StringComparison.Ordinal);
                    end = AddPrefixed(tokens, formula, i, end + 1, name, multi, external);
                }
                else
                {
                    tokens.Add(new FormulaToken(FormulaTokenKind.Other, i, end - i));
                }
            }
            else if (c == '#')
            {
                end = SkipError(formula, i);
                tokens.Add(new FormulaToken(end - i > 1 ? FormulaTokenKind.Error : FormulaTokenKind.Other, i, end - i));
            }
            else if (char.IsAsciiDigit(c) || (c == '.' && i + 1 < formula.Length && char.IsAsciiDigit(formula[i + 1])))
            {
                if (A1Area.TryParse(formula, i, out var area, out end) && area.Kind == A1AreaKind.Rows)
                {
                    tokens.Add(new FormulaToken(FormulaTokenKind.Reference, i, end - i) { Area = area });
                }
                else
                {
                    end = SkipNumber(formula, i);
                    tokens.Add(new FormulaToken(FormulaTokenKind.Number, i, end - i));
                }
            }
            else if (c == '$' || IsNameChar(c))
            {
                end = ScanNameOrReference(tokens, formula, i, lastBracketEnd);
            }
            else
            {
                end = i + 1;
                tokens.Add(new FormulaToken(FormulaTokenKind.Other, i, 1));
            }
            i = end;
        }
        return tokens;
    }

    private static int ScanNameOrReference(List<FormulaToken> tokens, string formula, int start, int lastBracketEnd)
    {
        if (A1Area.TryParse(formula, start, out var area, out var end))
        {
            tokens.Add(new FormulaToken(FormulaTokenKind.Reference, start, end - start) { Area = area });
            return end;
        }
        end = SkipName(formula, start);
        if (end == start)
        {
            // A lone '$' that does not start a reference.
            tokens.Add(new FormulaToken(FormulaTokenKind.Other, start, 1));
            return start + 1;
        }
        if (end < formula.Length && formula[end] == '!')
        {
            return AddPrefixed(tokens, formula, start, end + 1, formula[start..end], start == lastBracketEnd, start == lastBracketEnd);
        }
        if (end < formula.Length && formula[end] == ':')
        {
            // Sheet1:Sheet3!A1, a reference to several sheets.
            var secondEnd = SkipName(formula, end + 1);
            if (secondEnd > end + 1 && secondEnd < formula.Length && formula[secondEnd] == '!')
            {
                return AddPrefixed(tokens, formula, start, secondEnd + 1, formula[start..secondEnd], true, start == lastBracketEnd);
            }
        }
        var kind = end < formula.Length && formula[end] == '(' ? FormulaTokenKind.Function : FormulaTokenKind.Name;
        tokens.Add(new FormulaToken(kind, start, end - start));
        return end;
    }

    /// <summary>
    /// Adds the token for a sheet prefix running from <paramref name="start"/> to <paramref name="afterBang"/>: a reference when an A1
    /// area follows, otherwise the prefix alone as Other (as in <c>Sheet1!MyName</c> or <c>Sheet1!#REF!</c>). Both carry the prefix's
    /// length and sheet name.
    /// </summary>
    private static int AddPrefixed(
        List<FormulaToken> tokens, string formula, int start, int afterBang, string sheetName, bool multi, bool external)
    {
        if (A1Area.TryParse(formula, afterBang, out var area, out var end))
        {
            tokens.Add(new FormulaToken(FormulaTokenKind.Reference, start, end - start)
            {
                Area = area,
                PrefixLength = afterBang - start,
                SheetName = sheetName,
                IsMultiSheetOrExternal = multi,
                IsExternal = external,
            });
            return end;
        }
        tokens.Add(new FormulaToken(FormulaTokenKind.Other, start, afterBang - start)
        {
            PrefixLength = afterBang - start,
            SheetName = sheetName,
            IsMultiSheetOrExternal = multi,
            IsExternal = external,
        });
        return afterBang;
    }

    private static int SkipName(string formula, int start)
    {
        var i = start;
        while (i < formula.Length && IsNameChar(formula[i]))
        {
            i++;
        }
        return i;
    }

    /// <summary>Skips quoted text where the quote is escaped by doubling it. An unclosed quote runs to the end.</summary>
    private static int SkipQuoted(string formula, int start, char quote)
    {
        var i = start + 1;
        while (i < formula.Length)
        {
            if (formula[i] == quote)
            {
                if (i + 1 < formula.Length && formula[i + 1] == quote)
                {
                    i += 2;
                    continue;
                }
                return i + 1;
            }
            i++;
        }
        return i;
    }

    /// <summary>
    /// Skips nested square brackets, as in <c>[[#This Row],[Price]]</c>; inside them an apostrophe escapes the next character, as in
    /// <c>[Col']']</c>. An unclosed bracket runs to the end.
    /// </summary>
    private static int SkipBracketed(string formula, int start)
    {
        var depth = 0;
        var i = start;
        while (i < formula.Length)
        {
            var c = formula[i];
            if (c == '\'')
            {
                i += 2;
                continue;
            }
            if (c == '[')
            {
                depth++;
            }
            else if (c == ']' && --depth == 0)
            {
                return i + 1;
            }
            i++;
        }
        return formula.Length;
    }

    /// <summary>Skips <c>#</c> and the letters, digits, <c>/</c> and <c>_</c> after it, then a closing <c>!</c> or <c>?</c>.</summary>
    private static int SkipError(string formula, int start)
    {
        var i = start + 1;
        while (i < formula.Length && (char.IsAsciiLetterOrDigit(formula[i]) || formula[i] is '/' or '_'))
        {
            i++;
        }
        if (i > start + 1 && i < formula.Length && formula[i] is '!' or '?')
        {
            i++;
        }
        return i;
    }

    private static int SkipNumber(string formula, int start)
    {
        var i = start;
        while (i < formula.Length && (char.IsAsciiDigit(formula[i]) || formula[i] == '.'))
        {
            i++;
        }
        if (i < formula.Length && formula[i] is 'e' or 'E')
        {
            var j = i + 1;
            if (j < formula.Length && formula[j] is '+' or '-')
            {
                j++;
            }
            if (j < formula.Length && char.IsAsciiDigit(formula[j]))
            {
                i = j;
                while (i < formula.Length && char.IsAsciiDigit(formula[i]))
                {
                    i++;
                }
            }
        }
        return i;
    }
}
