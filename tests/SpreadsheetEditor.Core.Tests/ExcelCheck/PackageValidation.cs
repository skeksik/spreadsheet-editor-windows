using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>Validates an <c>.xlsx</c> package against the Open XML schemas and semantic rules with the Open XML SDK's validator.</summary>
internal static class PackageValidation
{
    /// <summary>
    /// The format version validated against: the newest the SDK knows, because Excel for Microsoft 365 is what opens the samples. The
    /// writer adds no extension elements of its own, so a newer version only widens what is accepted from a source package.
    /// </summary>
    public const FileFormatVersions Version = FileFormatVersions.Microsoft365;

    /// <summary>Every validation error, one line each: part, element path, error id and description; empty when the package is valid.</summary>
    public static List<string> Errors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, false);
        return new OpenXmlValidator(Version).Validate(document)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Id} {e.Description}")
            .ToList();
    }
}
