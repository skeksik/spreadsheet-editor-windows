using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Reads a workbook's shared string table with the SDK's SAX reader: one string per <c>si</c>, the text of its runs joined, without
/// phonetic (<c>rPh</c>) text.
/// </summary>
internal static class SharedStrings
{
    public static string[] Read(SharedStringTablePart? part, CancellationToken cancellationToken)
    {
        if (part is null)
        {
            return [];
        }
        var strings = new List<string>();
        StringBuilder? current = null;
        var inPhonetic = false;
        using var reader = OpenXmlReader.Create(part);
        while (reader.Read())
        {
            var type = reader.ElementType;
            if (reader.IsStartElement)
            {
                if (type == typeof(X.SharedStringItem))
                {
                    // A string is finished when the next starts (or the table ends), so an empty <si/> counts too.
                    if (current is not null)
                    {
                        strings.Add(current.ToString());
                    }
                    current = new StringBuilder();
                    inPhonetic = false;
                    if ((strings.Count & 0xFFF) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                }
                else if (type == typeof(X.PhoneticRun))
                {
                    inPhonetic = true;
                }
                else if (type == typeof(X.Run))
                {
                    inPhonetic = false;
                }
                else if (type == typeof(X.Text) && !inPhonetic && current is not null)
                {
                    current.Append(reader.GetText());
                }
            }
            else if (reader.IsEndElement && type == typeof(X.PhoneticRun))
            {
                inPhonetic = false;
            }
        }
        if (current is not null)
        {
            strings.Add(current.ToString());
        }
        return [.. strings];
    }
}
