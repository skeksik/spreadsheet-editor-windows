using System.Globalization;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Writes a workbook into a fresh copy of the package it was read from, or into a new empty package (docs/spec.md §5.1, §9.10, §9.12,
/// §9.21, §9.30–9.32). The source package itself is never changed, so saving twice gives the same file.
/// <list type="bullet">
/// <item>Each sheet read from a worksheet is written into it in place; a renamed sheet renames it; a duplicated sheet
/// (<see cref="SheetSource.IsCopy"/>) gets a copy of the whole worksheet; a new sheet gets a new worksheet.</item>
/// <item>Worksheets the app does not show (no table, chart sheets, dialog sheets) stay as they are, in their place; a worksheet with a
/// table that no sheet refers to any more was removed, and is deleted with its local defined names and the global ones that refer only
/// to it.</item>
/// <item>The other places in the sheet list follow the model's order.</item>
/// </list>
/// A rename or deletion that would leave formulas or defined names referring to a sheet name that no longer exists is refused until
/// formula references can be rewritten (docs/roadmap.md).
/// </summary>
internal sealed class WorkbookWriter
{
    private readonly Model.Workbook model;
    private readonly CultureInfo culture;
    private readonly CancellationToken cancellationToken;

    private WorkbookWriter(Model.Workbook model, CultureInfo culture, CancellationToken cancellationToken)
    {
        this.model = model;
        this.culture = culture;
        this.cancellationToken = cancellationToken;
    }

    public static byte[] Write(Model.Workbook model, CultureInfo culture, CancellationToken cancellationToken) =>
        new WorkbookWriter(model, culture, cancellationToken).Write();

    private byte[] Write()
    {
        if (model.Sheets.Count == 0)
        {
            throw SpreadsheetWriteException.Invalid("The workbook has no sheet; a workbook needs at least one.");
        }

        var package = model.Source?.Package.ToArray() ?? NewPackage.Create();
        using var sourceStream = new MemoryStream(package, writable: false);
        using var output = new MemoryStream();
        output.Write(package);
        output.Position = 0;

        SpreadsheetDocument? original = null;
        try
        {
            try
            {
                original = SpreadsheetDocument.Open(sourceStream, false);
            }
            catch (Exception ex) when (ex is not (OperationCanceledException or OutOfMemoryException))
            {
                throw new SpreadsheetWriteException(
                    SpreadsheetWriteError.SourceDamaged, $"The workbook the sheets were read from cannot be opened again: {ex.Message}", ex);
            }

            using (var document = SpreadsheetDocument.Open(output, true))
            {
                WriteDocument(original, document);
            }
        }
        finally
        {
            original?.Dispose();
        }

        FixTimestamps(output);
        return output.ToArray();
    }

    /// <summary>
    /// Gives every zip entry the same fixed time. The package stamps each entry with the time of the save (seen in this project's tests,
    /// 2026-11-03), which would make two saves of the same workbook differ (docs/spec.md §9.21). 1 January 1980 is taken as the earliest
    /// time a zip entry's DOS date can hold (not verified against the zip format's specification).
    /// </summary>
    private static void FixTimestamps(MemoryStream package)
    {
        var fixedTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true);
        foreach (var entry in archive.Entries)
        {
            entry.LastWriteTime = fixedTime;
        }
    }

    private void WriteDocument(SpreadsheetDocument original, SpreadsheetDocument document)
    {
        var sourceBook = original.WorkbookPart ?? throw Damaged("The package has no workbook part.");
        var workbookPart = document.WorkbookPart ?? throw Damaged("The package has no workbook part.");
        var root = workbookPart.Workbook ?? throw Damaged("The workbook part is empty.");
        var sheetList = root.Sheets ??= new X.Sheets();
        var system = root.WorkbookProperties?.Date1904?.Value == true ? DateSystem.Date1904 : DateSystem.Date1900;
        var reader = new WorksheetCells(
            SharedStrings.Read(sourceBook.SharedStringTablePart, cancellationToken),
            NumberFormatTable.Read(sourceBook.WorkbookStylesPart?.Stylesheet),
            system,
            culture);

        // The worksheets as listed, and which model sheet writes into each one.
        var entries = sheetList.Elements<X.Sheet>().ToList();
        var byId = entries.Where(e => e.Id?.Value is not null).ToDictionary(e => e.Id!.Value!, StringComparer.Ordinal);
        var plans = new List<SheetPlan>();
        var claimed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sheet in model.Sheets)
        {
            if (sheet.Source is not { } source)
            {
                plans.Add(new SheetPlan(sheet, null, null));
                continue;
            }

            if (!byId.TryGetValue(source.RelationshipId, out var entry) || sourceBook.GetPartById(source.RelationshipId) is not WorksheetPart)
            {
                throw Damaged($"The worksheet the sheet \"{sheet.Name}\" was read from is missing.");
            }

            // A duplicated sheet copies its worksheet; so does a second sheet claiming the same worksheet.
            var inPlace = !source.IsCopy && claimed.Add(source.RelationshipId);
            plans.Add(new SheetPlan(sheet, inPlace ? entry : null, source.RelationshipId));
        }

        // Worksheets no sheet writes into: kept when the app does not show them, deleted when they held a table (the tab was removed).
        var kept = new List<X.Sheet>();
        var deleted = new List<X.Sheet>();
        foreach (var entry in entries)
        {
            if (entry.Id?.Value is { } id && claimed.Contains(id))
            {
                continue;
            }

            var part = entry.Id?.Value is { } relationship ? sourceBook.GetPartById(relationship) : null;
            if (part is WorksheetPart worksheet && HasTable(worksheet, reader))
            {
                deleted.Add(entry);
            }
            else
            {
                kept.Add(entry);
            }
        }

        CheckNames(plans, kept);
        CheckReferences(original, plans, deleted);

        var sources = new Dictionary<string, SourceWorksheet>(StringComparer.Ordinal);
        SourceWorksheet SourceOf(string relationshipId)
        {
            if (!sources.TryGetValue(relationshipId, out var found))
            {
                found = new SourceWorksheet((WorksheetPart)sourceBook.GetPartById(relationshipId), reader, cancellationToken);
                sources[relationshipId] = found;
            }

            return found;
        }

        var oldIndex = entries.Select((e, i) => (e, i)).ToDictionary(p => p.e, p => p.i);
        var definedNames = root.DefinedNames;
        var localNames = definedNames?.Elements<X.DefinedName>().Where(n => n.LocalSheetId is not null).ToList() ?? [];

        // Copies first, while every worksheet still holds what was read.
        var nextSheetId = entries.Count == 0 ? 1u : entries.Max(e => e.SheetId?.Value ?? 0) + 1;
        var copiedNames = new List<(X.DefinedName Name, SheetPlan Plan)>();
        foreach (var plan in plans.Where(p => p.Entry is null))
        {
            cancellationToken.ThrowIfCancellationRequested();
            WorksheetPart part;
            X.Sheet entry;
            if (plan.SourceId is { } sourceId)
            {
                var originalEntry = byId[sourceId];
                part = WorksheetCopier.Copy(workbookPart, (WorksheetPart)workbookPart.GetPartById(sourceId), plan.Sheet.Name);
                entry = (X.Sheet)originalEntry.CloneNode(true);
                foreach (var name in localNames.Where(n => n.LocalSheetId!.Value == oldIndex[originalEntry]))
                {
                    var copy = (X.DefinedName)name.CloneNode(true);
                    copy.Text = SheetReferences.Replace(name.Text, originalEntry.Name?.Value ?? "", plan.Sheet.Name);
                    copiedNames.Add((copy, plan));
                }
            }
            else
            {
                part = workbookPart.AddNewPart<WorksheetPart>(RelationshipIds.Next(workbookPart));
                part.Worksheet = new X.Worksheet(new X.SheetData());
                entry = new X.Sheet();
            }

            entry.Id = workbookPart.GetIdOfPart(part);
            entry.SheetId = nextSheetId++;
            plan.Entry = entry;
        }

        foreach (var entry in deleted)
        {
            DeleteWorksheet(workbookPart, root, entry, oldIndex[entry]);
        }

        // The new sheet list: worksheets the app does not show keep their places; the model's sheets fill the others in tab order.
        var order = new List<X.Sheet>();
        var next = 0;
        foreach (var entry in entries.Where(e => !deleted.Contains(e)))
        {
            order.Add(kept.Contains(entry) ? entry : plans[next++].Entry!);
        }

        while (next < plans.Count)
        {
            order.Add(plans[next++].Entry!);
        }

        foreach (var plan in plans)
        {
            plan.Entry!.Name = plan.Sheet.Name;
        }

        Reorder(root, sheetList, entries, order, oldIndex, copiedNames);
        CheckVisible(order);

        var sourceStyles = sourceBook.WorkbookStylesPart?.Stylesheet;
        var styles = new StyleWriter(
            workbookPart,
            StyleTable.Read(sourceStyles, sourceBook.ThemePart),
            sourceStyles?.CellFormats?.Elements<X.CellFormat>().Count() ?? 0);
        var context = new WriteContext(culture, system, new SharedStringWriter(workbookPart), styles);
        foreach (var plan in plans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var part = (WorksheetPart)workbookPart.GetPartById(plan.Entry!.Id!.Value!);
            var source = plan.SourceId is { } id ? SourceOf(id) : SourceWorksheet.Empty;
            WorksheetWriter.Write(plan.Sheet, part.Worksheet!, source, context);
        }

        context.Strings.Finish();
        if (deleted.Count > 0)
        {
            context.FormulaChanged = true;
        }

        Recalculate(workbookPart, root, context);
    }

    private static bool HasTable(WorksheetPart part, WorksheetCells reader) =>
        TableDetector.Detect(reader.Read(part, CancellationToken.None), []) is not null;

    /// <summary>Sheet names are unique without regard to case, among the model's sheets and the worksheets kept as they are.</summary>
    private static void CheckNames(List<SheetPlan> plans, List<X.Sheet> kept)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in kept.Select(k => k.Name?.Value ?? "").Concat(plans.Select(p => p.Sheet.Name)))
        {
            if (!names.Add(name))
            {
                throw SpreadsheetWriteException.Invalid($"Two sheets in the workbook are named \"{name}\"; sheet names must be unique.");
            }
        }
    }

    /// <summary>
    /// Refuses a rename or deletion that would leave a reference to a sheet name that no longer exists: in a defined name, a cell formula,
    /// a chart or any other part (a conservative text check, <see cref="SheetReferences"/>). Defined names that are local to a deleted
    /// worksheet, or that refer only to it, are deleted with it instead.
    /// </summary>
    private static void CheckReferences(SpreadsheetDocument original, List<SheetPlan> plans, List<X.Sheet> deleted)
    {
        var gone = new List<(string Name, string? SkipPart)>();
        foreach (var plan in plans)
        {
            if (plan.Entry is { } entry && entry.Name?.Value is { } old && !string.Equals(old, plan.Sheet.Name, StringComparison.Ordinal))
            {
                gone.Add((old, null));
            }
        }

        var book = original.WorkbookPart!;
        var entries = book.Workbook!.Sheets!.Elements<X.Sheet>().ToList();
        var deletedIndexes = deleted.Select(e => entries.IndexOf(e)).ToHashSet();
        foreach (var entry in deleted)
        {
            gone.Add((entry.Name?.Value ?? "", entry.Id?.Value is { } id ? book.GetPartById(id).Uri.OriginalString : null));
        }

        if (gone.Count == 0)
        {
            return;
        }

        foreach (var name in book.Workbook.DefinedNames?.Elements<X.DefinedName>() ?? [])
        {
            if (name.LocalSheetId?.Value is { } local && deletedIndexes.Contains((int)local))
            {
                continue;
            }

            var refs = SheetReferences.Find(name.Text, skipStrings: true);
            foreach (var (sheetName, _) in gone)
            {
                var deletedName = deleted.Exists(d => string.Equals(d.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                var onlyThis = refs.Count > 0 && refs.TrueForAll(r => string.Equals(r, sheetName, StringComparison.OrdinalIgnoreCase));
                if (deletedName && onlyThis)
                {
                    continue;
                }

                if (refs.Exists(r => string.Equals(r, sheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    throw References(sheetName, $"the defined name \"{name.Name?.Value}\"");
                }
            }
        }

        foreach (var part in AllParts(original))
        {
            if (part == book || part is SharedStringTablePart || !IsXml(part))
            {
                continue;
            }

            string text;
            using (var reader = new StreamReader(part.GetStream()))
            {
                text = Unescape(reader.ReadToEnd());
            }

            foreach (var (sheetName, skipPart) in gone)
            {
                if (part.Uri.OriginalString == skipPart)
                {
                    continue;
                }

                // Pivot caches name their source sheet in an attribute, not as a reference.
                if (SheetReferences.Mentions(text, sheetName, skipStrings: false)
                    || text.Contains("sheet=\"" + sheetName + "\"", StringComparison.OrdinalIgnoreCase))
                {
                    throw References(sheetName, part.Uri.OriginalString);
                }
            }
        }
    }

    private static SpreadsheetWriteException References(string sheetName, string where) =>
        SpreadsheetWriteException.NotSupported(
            $"The sheet \"{sheetName}\" was renamed or removed, but {where} still refers to it by name; references cannot be rewritten yet.");

    private static IEnumerable<OpenXmlPart> AllParts(SpreadsheetDocument document)
    {
        var seen = new HashSet<OpenXmlPart>();
        var stack = new Stack<OpenXmlPartContainer>();
        stack.Push(document);
        while (stack.Count > 0)
        {
            foreach (var pair in stack.Pop().Parts)
            {
                if (seen.Add(pair.OpenXmlPart))
                {
                    stack.Push(pair.OpenXmlPart);
                    yield return pair.OpenXmlPart;
                }
            }
        }
    }

    private static bool IsXml(OpenXmlPart part) => part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase);

    private static string Unescape(string xml) => xml
        .Replace("&apos;", "'", StringComparison.Ordinal)
        .Replace("&quot;", "\"", StringComparison.Ordinal)
        .Replace("&lt;", "<", StringComparison.Ordinal)
        .Replace("&gt;", ">", StringComparison.Ordinal)
        .Replace("&amp;", "&", StringComparison.Ordinal);

    private static void DeleteWorksheet(WorkbookPart workbookPart, X.Workbook root, X.Sheet entry, int index)
    {
        if (root.DefinedNames is { } names)
        {
            var sheetName = entry.Name?.Value ?? "";
            foreach (var name in names.Elements<X.DefinedName>().ToList())
            {
                var local = name.LocalSheetId?.Value is { } l && l == index;
                var refs = SheetReferences.Find(name.Text, skipStrings: true);
                var onlyThis = name.LocalSheetId is null && refs.Count > 0
                    && refs.TrueForAll(r => string.Equals(r, sheetName, StringComparison.OrdinalIgnoreCase));
                if (local || onlyThis)
                {
                    name.Remove();
                }
            }

            if (!names.HasChildren)
            {
                names.Remove();
            }
        }

        if (entry.Id?.Value is { } id)
        {
            workbookPart.DeletePart(id);
        }
    }

    /// <summary>
    /// Puts the sheet list in its new order and moves what refers to sheets by position with it: local defined names
    /// (<c>localSheetId</c>) and the workbook views' active and first tabs.
    /// </summary>
    private static void Reorder(
        X.Workbook root,
        X.Sheets sheetList,
        List<X.Sheet> entries,
        List<X.Sheet> order,
        Dictionary<X.Sheet, int> oldIndex,
        List<(X.DefinedName Name, SheetPlan Plan)> copiedNames)
    {
        var newIndex = order.Select((e, i) => (e, i)).ToDictionary(p => p.e, p => p.i);
        var map = entries.ToDictionary(e => oldIndex[e], e => newIndex.TryGetValue(e, out var i) ? i : -1);

        sheetList.RemoveAllChildren();
        foreach (var entry in order)
        {
            if (entry.Parent is not null)
            {
                entry.Remove();
            }

            sheetList.AppendChild(entry);
        }

        if (root.DefinedNames is { } names)
        {
            foreach (var name in names.Elements<X.DefinedName>())
            {
                if (name.LocalSheetId?.Value is { } local && map.TryGetValue((int)local, out var moved) && moved >= 0)
                {
                    name.LocalSheetId = (uint)moved;
                }
            }
        }

        if (copiedNames.Count > 0)
        {
            var list = root.DefinedNames ?? InsertDefinedNames(root);
            foreach (var (name, plan) in copiedNames)
            {
                name.LocalSheetId = (uint)newIndex[plan.Entry!];
                list.AppendChild(name);
            }
        }

        foreach (var view in root.BookViews?.Elements<X.WorkbookView>() ?? [])
        {
            view.ActiveTab = Remap(view.ActiveTab?.Value, map, order);
            view.FirstSheet = Remap(view.FirstSheet?.Value, map, order);
        }
    }

    private static X.DefinedNames InsertDefinedNames(X.Workbook root)
    {
        var names = new X.DefinedNames();
        root.AddChild(names, throwOnError: false);
        return names;
    }

    /// <summary>The new index of a tab, or the first visible sheet when that tab was deleted; null stays null.</summary>
    private static uint? Remap(uint? index, Dictionary<int, int> map, List<X.Sheet> order)
    {
        if (index is not { } old)
        {
            return null;
        }

        if (map.TryGetValue((int)old, out var moved) && moved >= 0)
        {
            return (uint)moved;
        }

        var visible = order.FindIndex(e => e.State is null || e.State.Value == X.SheetStateValues.Visible);
        return (uint)Math.Max(0, visible);
    }

    private static void CheckVisible(List<X.Sheet> order)
    {
        if (!order.Exists(e => e.State is null || e.State.Value == X.SheetStateValues.Visible))
        {
            throw SpreadsheetWriteException.Invalid("Every sheet in the workbook is hidden; a workbook needs at least one visible sheet.");
        }
    }

    /// <summary>
    /// After values changed, the workbook asks for a full calculation when it is opened, so formulas show results for the new values; a
    /// stale calculation chain is dropped. Microsoft Learn, "SheetCalculationProperties Class" (quoting ISO/IEC 29500-1 1st edition, read
    /// 2026-11-03), on the same attribute of the worksheet's element: "fullCalcOnLoad … Indicates whether the application should do a full
    /// calculate on load due to contents on this sheet. After load and successful calc, the application shall set this value to false. Set
    /// this to true when the application should calculate the workbook on load." The workbook's <c>calcPr</c> carries the same attribute
    /// ("CalculationProperties Class", same date). That a spreadsheet program rebuilds a missing calculation chain is not verified.
    /// </summary>
    private static void Recalculate(WorkbookPart workbookPart, X.Workbook root, WriteContext context)
    {
        if (context.ValueChanged)
        {
            var properties = root.CalculationProperties;
            if (properties is null)
            {
                properties = new X.CalculationProperties();
                root.AddChild(properties, throwOnError: false);
            }

            properties.FullCalculationOnLoad = true;
        }

        if (context.FormulaChanged && workbookPart.CalculationChainPart is { } chain)
        {
            workbookPart.DeletePart(chain);
        }
    }

    private static SpreadsheetWriteException Damaged(string message) =>
        new(SpreadsheetWriteError.SourceDamaged, $"The workbook the sheets were read from is damaged: {message}");

    private sealed class SheetPlan(Sheet sheet, X.Sheet? entry, string? sourceId)
    {
        public Sheet Sheet { get; } = sheet;

        /// <summary>The worksheet's entry in the sheet list: the one read (in place), or a new one (copy or new sheet).</summary>
        public X.Sheet? Entry { get; set; } = entry;

        /// <summary>The relationship id of the worksheet in the source the sheet was read from, or null for a new sheet.</summary>
        public string? SourceId { get; } = sourceId;
    }
}
