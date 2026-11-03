using System.Globalization;
using DocumentFormat.OpenXml.Packaging;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Relationship ids for the parts the writer adds. The SDK makes a random id when none is given, which would make two saves of the same
/// workbook differ (docs/spec.md §9.21); the writer gives every new part the first free id of the form <c>rId{n}</c> instead.
/// </summary>
internal static class RelationshipIds
{
    public static string Next(OpenXmlPartContainer container)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in container.Parts)
        {
            used.Add(pair.RelationshipId);
        }

        foreach (var relationship in container.ExternalRelationships)
        {
            used.Add(relationship.Id);
        }

        foreach (var relationship in container.HyperlinkRelationships)
        {
            used.Add(relationship.Id);
        }

        foreach (var relationship in container.DataPartReferenceRelationships)
        {
            used.Add(relationship.Id);
        }

        for (var n = 1; ; n++)
        {
            var id = "rId" + n.ToString(CultureInfo.InvariantCulture);
            if (!used.Contains(id))
            {
                return id;
            }
        }
    }
}
