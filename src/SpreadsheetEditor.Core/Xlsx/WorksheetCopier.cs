using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Copies a worksheet part with the parts it uses, for a duplicated sheet (docs/spec.md §9.30): drawings and their charts, comments and
/// their VML, printer settings and tables are copied; pictures are shared with the original, as their bytes do not change. Tables get a new
/// id and a unique name. A worksheet that uses any other kind of part (pivot tables, slicers, query tables, threaded comments, ...) is
/// refused, as copying those needs more than their bytes.
/// </summary>
internal static class WorksheetCopier
{
    public static WorksheetPart Copy(WorkbookPart workbookPart, WorksheetPart original, string sheetName)
    {
        var copy = workbookPart.AddNewPart<WorksheetPart>(RelationshipIds.Next(workbookPart));
        CopyContent(original, copy, sheetName);
        var worksheet = copy.Worksheet!;
        // Only one tab is selected; the copy is not.
        foreach (var view in worksheet.Descendants<X.SheetView>())
        {
            view.TabSelected = null;
        }

        RenameTables(workbookPart, copy);
        return copy;
    }

    private static void CopyContent(OpenXmlPart from, OpenXmlPart to, string sheetName)
    {
        using (var stream = from.GetStream())
        {
            to.FeedData(stream);
        }

        foreach (var relationship in from.ExternalRelationships)
        {
            to.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
        }

        foreach (var relationship in from.HyperlinkRelationships)
        {
            to.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
        }

        if (from.DataPartReferenceRelationships.Any())
        {
            throw Unsupported(sheetName, "media");
        }

        foreach (var pair in from.Parts)
        {
            var id = pair.RelationshipId;
            switch (pair.OpenXmlPart)
            {
                case ImagePart image:
                    to.AddPart(image, id);
                    break;
                case DrawingsPart part:
                    CopyContent(part, to.AddNewPart<DrawingsPart>(id), sheetName);
                    break;
                case ChartPart part:
                    CopyContent(part, to.AddNewPart<ChartPart>(id), sheetName);
                    break;
                case ChartStylePart part:
                    CopyContent(part, to.AddNewPart<ChartStylePart>(id), sheetName);
                    break;
                case ChartColorStylePart part:
                    CopyContent(part, to.AddNewPart<ChartColorStylePart>(id), sheetName);
                    break;
                case ChartDrawingPart part:
                    CopyContent(part, to.AddNewPart<ChartDrawingPart>(id), sheetName);
                    break;
                case ThemeOverridePart part:
                    CopyContent(part, to.AddNewPart<ThemeOverridePart>(id), sheetName);
                    break;
                case EmbeddedPackagePart part:
                    CopyContent(part, to.AddNewPart<EmbeddedPackagePart>(part.ContentType, id), sheetName);
                    break;
                case VmlDrawingPart part:
                    CopyContent(part, to.AddNewPart<VmlDrawingPart>(id), sheetName);
                    break;
                case LegacyDiagramTextPart part:
                    CopyContent(part, to.AddNewPart<LegacyDiagramTextPart>(id), sheetName);
                    break;
                case WorksheetCommentsPart part:
                    CopyContent(part, to.AddNewPart<WorksheetCommentsPart>(id), sheetName);
                    break;
                case SpreadsheetPrinterSettingsPart part:
                    CopyContent(part, to.AddNewPart<SpreadsheetPrinterSettingsPart>(id), sheetName);
                    break;
                case TableDefinitionPart part:
                    CopyContent(part, to.AddNewPart<TableDefinitionPart>(id), sheetName);
                    break;
                case var other:
                    throw Unsupported(sheetName, other.ContentType);
            }
        }
    }

    /// <summary>Gives each table of the copy an id and a name no other table has (a table's name is unique in the workbook).</summary>
    private static void RenameTables(WorkbookPart workbookPart, WorksheetPart copy)
    {
        var copied = copy.TableDefinitionParts.ToList();
        if (copied.Count == 0)
        {
            return;
        }

        var others = workbookPart.WorksheetParts
            .Where(p => p != copy)
            .SelectMany(p => p.TableDefinitionParts)
            .Select(p => p.Table)
            .OfType<X.Table>()
            .ToList();
        var ids = others.Select(t => t.Id?.Value ?? 0).ToHashSet();
        var names = others.SelectMany(t => new[] { t.Name?.Value, t.DisplayName?.Value }).OfType<string>()
            .Concat(workbookPart.Workbook?.DefinedNames?.Elements<X.DefinedName>().Select(n => n.Name?.Value).OfType<string>() ?? [])
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var nextId = ids.Count == 0 ? 1u : ids.Max() + 1;
        foreach (var part in copied)
        {
            var table = part.Table!;
            table.Id = nextId++;
            var baseName = table.DisplayName?.Value ?? table.Name?.Value ?? "Table";
            var n = 2;
            var name = baseName + "_" + n;
            while (names.Contains(name))
            {
                name = baseName + "_" + ++n;
            }

            names.Add(name);
            table.Name = name;
            table.DisplayName = name;
        }
    }

    private static SpreadsheetWriteException Unsupported(string sheetName, string what) =>
        SpreadsheetWriteException.NotSupported($"The sheet \"{sheetName}\" holds content of a kind that cannot be copied yet ({what}).");
}
