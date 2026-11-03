namespace SpreadsheetEditor.Core.Formulas;

/// <summary>The kinds of text <see cref="FormulaTokenizer"/> tells apart.</summary>
internal enum FormulaTokenKind
{
    /// <summary>Operators, separators, whitespace, a sheet prefix not followed by a reference, and anything else.</summary>
    Other,

    /// <summary>A string literal in double quotes, <c>""</c> inside it standing for one quote.</summary>
    String,

    /// <summary>Text in square brackets: a structured (table) reference part or an external workbook index.</summary>
    Bracketed,

    /// <summary>An error literal such as <c>#REF!</c> or <c>#N/A</c>.</summary>
    Error,

    /// <summary>A number literal.</summary>
    Number,

    /// <summary>A function name: a name directly followed by <c>(</c>.</summary>
    Function,

    /// <summary>A defined name, a table name, <c>TRUE</c> or <c>FALSE</c>: a name that is not a reference or a function.</summary>
    Name,

    /// <summary>An A1 reference with its optional sheet prefix.</summary>
    Reference,
}
