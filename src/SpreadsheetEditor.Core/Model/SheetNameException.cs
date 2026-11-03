namespace SpreadsheetEditor.Core.Model;

/// <summary>A sheet name breaks a worksheet name rule (docs/spec.md §9.39); <see cref="Check"/> says which, for a localised message.</summary>
public sealed class SheetNameException : ArgumentException
{
    public SheetNameException()
    {
    }

    public SheetNameException(string message)
        : base(message)
    {
    }

    public SheetNameException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public SheetNameException(string name, SheetNameCheck check)
        : base($"The sheet name '{name}' is not allowed: {check.Error}.", nameof(name))
    {
        Check = check;
    }

    /// <summary>The rule the name breaks.</summary>
    public SheetNameCheck Check { get; }
}
