namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>What a colour in the workbook's XML resolved to (docs/spec.md §4.4, §9.5, §9.6).</summary>
internal enum ColorKind
{
    /// <summary>The colour cannot be resolved (an index out of range, a gradient, a colour without attributes): leave it as it is.</summary>
    Unresolved,

    /// <summary>An explicit, fully transparent colour (<c>00000000</c>): no colour.</summary>
    None,

    /// <summary>An opaque colour.</summary>
    Rgb,

    /// <summary>The automatic colour: <c>auto="1"</c>, or the system foreground and background (indexed 64 and 65).</summary>
    Automatic,
}
