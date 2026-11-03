using System.Buffers;
using System.Globalization;

namespace SpreadsheetEditor.Core.Model;

/// <summary>
/// The rules for worksheet names (docs/spec.md §9.39) and the names given to new and duplicated sheets (§2.3).
/// </summary>
/// <remarks>
/// Sources, read 2026-11-03:
/// <list type="bullet">
/// <item>Microsoft Support, "Rename a worksheet"
/// (https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9): names cannot "Be blank",
/// "Contain more than 31 characters", "Contain any of the following characters: / \ ? * : [ ]", or "Begin or end with an apostrophe
/// ('), but they can be used in between text or numbers in a name", and cannot be named "History": "This is a reserved word Excel uses
/// internally."</item>
/// <item>[MS-OI29500] 2.1.578, Part 1 §18.2.19 sheet
/// (https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ebf12ea5-2bb4-4af5-ab26-563f22d3f895): "Office requires
/// that the name attribute be less than or equal to 31 characters in length" and "that sheet names are case-insensitive unique in any
/// locale."</item>
/// <item>[MS-OI29500] 2.1.792, Part 1 §18.10.1.79 rangeSet
/// (https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a9b7cdf8-64e5-47c9-950d-ebe253c5c82e), on a sheet-name
/// attribute: it "shall not be empty and its length shall not exceed 31 characters", "shall not begin or end with the apostrophe
/// (0x0027) character" and shall not contain Null (0x0000), End Of Text (0x0003), * / : ? [ \ ].</item>
/// <item>.NET, "Best practices for comparing strings"
/// (https://learn.microsoft.com/en-us/dotnet/standard/base-types/best-practices-strings): "An OrdinalIgnoreCase comparer still operates
/// on a char-by-char basis, but it eliminates case differences while performing the operation", for example 'á' and 'Á' compare
/// equal; ordinal comparisons do not depend on the culture.</item>
/// </list>
/// Not verified: that Excel counts a name's length in UTF-16 code units (this class does, which is never more lenient than counting
/// code points); that "History" is reserved ignoring case (this class rejects "history" too); and that Excel's case folding equals
/// <see cref="StringComparison.OrdinalIgnoreCase"/> for every script (Microsoft says only "in any locale").
/// </remarks>
public static class SheetNames
{
    /// <summary>The longest worksheet name, in UTF-16 code units.</summary>
    public const int MaxLength = 31;

    /// <summary>The name Excel reserves for its own use.</summary>
    public const string ReservedName = "History";

    /// <summary>The suffix of a duplicated sheet's name, never translated (§2.3).</summary>
    private const string CopyWord = "Copy";

    /// <summary>How sheet names compare: ignoring case, the same in every locale.</summary>
    public static StringComparer Comparer { get; } = StringComparer.OrdinalIgnoreCase;

    private static readonly SearchValues<char> InvalidCharacters = SearchValues.Create("\u0000\u0003:\\/?*[]");

    /// <summary>
    /// Checks <paramref name="name"/> against the worksheet name rules and the names already in use.
    /// </summary>
    /// <param name="name">The proposed name.</param>
    /// <param name="existingNames">The names of the workbook's sheets.</param>
    /// <param name="renamedName">
    /// When renaming, the sheet's current name: it is skipped once in the duplicate check, so a sheet can be renamed to a change of case.
    /// </param>
    public static SheetNameCheck Check(string name, IEnumerable<string> existingNames, string? renamedName = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(existingNames);

        if (name.Length == 0)
        {
            return new(SheetNameError.Empty);
        }

        if (name.Length > MaxLength)
        {
            return new(SheetNameError.TooLong);
        }

        var invalid = name.AsSpan().IndexOfAny(InvalidCharacters);
        if (invalid >= 0)
        {
            return new(SheetNameError.InvalidCharacter, name[invalid]);
        }

        if (name[0] == '\'')
        {
            return new(SheetNameError.LeadingApostrophe);
        }

        if (name[^1] == '\'')
        {
            return new(SheetNameError.TrailingApostrophe);
        }

        if (Comparer.Equals(name, ReservedName))
        {
            return new(SheetNameError.Reserved);
        }

        var skipped = renamedName is null;
        foreach (var existing in existingNames)
        {
            if (!skipped && string.Equals(existing, renamedName, StringComparison.Ordinal))
            {
                skipped = true;
                continue;
            }

            if (Comparer.Equals(existing, name))
            {
                return new(SheetNameError.Duplicate);
            }
        }

        return SheetNameCheck.Valid;
    }

    /// <summary>
    /// A name for a new sheet: <paramref name="baseName"/> followed by the smallest number from 1 that makes it unused, ignoring case
    /// ("Sheet1", "Sheet2", ...). The base is shortened, never splitting a surrogate pair, so the name fits in <see cref="MaxLength"/>.
    /// </summary>
    /// <param name="baseName">The localised word for a sheet, supplied by the app.</param>
    /// <param name="existingNames">The names of the workbook's sheets.</param>
    public static string NewSheetName(string baseName, IEnumerable<string> existingNames)
    {
        ArgumentNullException.ThrowIfNull(baseName);
        var used = ToSet(existingNames);
        for (var n = 1; ; n++)
        {
            var name = Fit(baseName, n.ToString(CultureInfo.InvariantCulture));
            if (!used.Contains(name))
            {
                return name;
            }
        }
    }

    /// <summary>
    /// A name for a copy of the sheet named <paramref name="sourceName"/>: "&lt;name&gt; (Copy)", or when that is used, ignoring case,
    /// "&lt;name&gt; (Copy 1)", "&lt;name&gt; (Copy 2)", ... (§2.3). The base is shortened, never splitting a surrogate pair, so the name
    /// fits in <see cref="MaxLength"/> (§9.39).
    /// </summary>
    /// <param name="sourceName">The name of the sheet being duplicated.</param>
    /// <param name="existingNames">The names of the workbook's sheets.</param>
    public static string CopyName(string sourceName, IEnumerable<string> existingNames)
    {
        ArgumentNullException.ThrowIfNull(sourceName);
        var used = ToSet(existingNames);
        var name = Fit(sourceName, $" ({CopyWord})");
        for (var n = 1; used.Contains(name); n++)
        {
            name = Fit(sourceName, string.Create(CultureInfo.InvariantCulture, $" ({CopyWord} {n})"));
        }

        return name;
    }

    /// <summary>
    /// <paramref name="baseName"/> followed by <paramref name="suffix"/>, with the base cut so the whole fits in <see cref="MaxLength"/>.
    /// A cut that would leave half of a surrogate pair drops the whole pair.
    /// </summary>
    internal static string Fit(string baseName, string suffix)
    {
        var room = Math.Max(0, MaxLength - suffix.Length);
        if (baseName.Length > room)
        {
            var cut = room;
            if (cut > 0 && char.IsHighSurrogate(baseName[cut - 1]))
            {
                cut--;
            }

            baseName = baseName[..cut];
        }

        return baseName + suffix;
    }

    private static HashSet<string> ToSet(IEnumerable<string> existingNames)
    {
        ArgumentNullException.ThrowIfNull(existingNames);
        return new HashSet<string>(existingNames, Comparer);
    }
}
