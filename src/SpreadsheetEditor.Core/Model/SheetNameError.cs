namespace SpreadsheetEditor.Core.Model;

/// <summary>Why a worksheet name is not allowed (docs/spec.md §9.39). The app turns each value into a localised message.</summary>
public enum SheetNameError
{
    /// <summary>The name is allowed.</summary>
    None,

    /// <summary>The name has no characters.</summary>
    Empty,

    /// <summary>The name is longer than <see cref="SheetNames.MaxLength"/> characters.</summary>
    TooLong,

    /// <summary>The name contains a character worksheet names cannot hold; <see cref="SheetNameCheck.Character"/> says which.</summary>
    InvalidCharacter,

    /// <summary>The name begins with an apostrophe (').</summary>
    LeadingApostrophe,

    /// <summary>The name ends with an apostrophe (').</summary>
    TrailingApostrophe,

    /// <summary>The name is one the spreadsheet program reserves (<see cref="SheetNames.ReservedName"/>).</summary>
    Reserved,

    /// <summary>Another sheet already has this name, ignoring case.</summary>
    Duplicate,
}
