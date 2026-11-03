namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>How a test picture is placed in a worksheet's drawing.</summary>
public enum PictureAnchor
{
    /// <summary><c>xdr:twoCellAnchor</c> holding <c>xdr:pic</c>.</summary>
    TwoCell,

    /// <summary><c>xdr:oneCellAnchor</c> holding <c>xdr:pic</c>.</summary>
    OneCell,

    /// <summary><c>xdr:absoluteAnchor</c>: no cell.</summary>
    Absolute,

    /// <summary>A <c>xdr:twoCellAnchor</c> holding a group (<c>xdr:grpSp</c>) with the picture inside.</summary>
    Grouped,

    /// <summary>A <c>xdr:twoCellAnchor</c> whose blip links to an external file (<c>r:link</c>) instead of embedding it.</summary>
    Linked,
}
