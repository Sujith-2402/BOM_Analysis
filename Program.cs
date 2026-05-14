using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

internal static class Program
{
    private static readonly string[] Headers =
    [
        "ParentPath", "Parent Filename", "Parent FileExt", "Parent Rev", "2D Orphan", "Multiple 3D",
        "Child Path", "Child Filename", "Child FileExt", "Child Rev", "Chilt Qty", "In Location Status",
        "With Filename Status"
    ];

    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.WriteLine("BOM Analysis Report Generator");
        Console.WriteLine("-----------------------------");

        try
        {
            var inputListPath = args.Length > 0 ? args[0] : PromptForExistingFile("Enter 1st input file path");
            var relationPath = args.Length > 1 ? args[1] : PromptForExistingFile("Enter 2nd input file path");
            var outputFolder = args.Length > 2 ? args[2] : PromptForOutputFolder("Enter output folder path");
            Directory.CreateDirectory(outputFolder);

            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            var logPath = Path.Combine(outputFolder, $"BOM_Analysis_Log_{timestamp}.txt");
            using var log = new StreamWriter(logPath, false, new UTF8Encoding(false));

            Log(log, "Started BOM Analysis.");
            Log(log, $"Input file 1: {inputListPath}");
            Log(log, $"Input file 2: {relationPath}");
            Log(log, $"Output folder: {outputFolder}");

            var locations = InputReader.ReadRows(inputListPath)
                .SelectMany(row => row)
                .Select(Clean)
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var exactLocations = new HashSet<string>(locations, StringComparer.OrdinalIgnoreCase);
            var fileNames = new HashSet<string>(locations.Select(Path.GetFileName).Where(IsPresent)!, StringComparer.OrdinalIgnoreCase);

            var relationRows = InputReader.ReadRows(relationPath).ToList();
            var parentSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var allRows = new List<string[]>();
            var orphanRows = new List<string[]>();
            var multiRows = new List<string[]>();
            var parentOnlyRows = 0;

            foreach (var sourceRow in relationRows)
            {
                var parts = sourceRow
                    .SelectMany(cell => cell.Split('|', StringSplitOptions.TrimEntries))
                    .Select(Clean)
                    .Where(IsPresent)
                    .ToList();

                if (parts.Count == 0)
                {
                    continue;
                }

                var parent = parts[0];
                var children = parts.Skip(1).ToList();
                parentSet.Add(parent);

                if (children.Count == 0)
                {
                    parentOnlyRows++;
                    AddRow(parent, child: "", childCount: 0, exactLocations, fileNames, allRows, orphanRows);
                    Log(log, $"Skipped row with no child: {parent}");
                    continue;
                }

                foreach (var child in children)
                {
                    var row = CreateReportRow(parent, child, children.Count, exactLocations, fileNames);
                    allRows.Add(row);
                    if (children.Count > 1)
                    {
                        multiRows.Add(row);
                    }
                }
            }

            foreach (var drawing in locations.Where(path => Path.GetExtension(path).Equals(".SLDDRW", StringComparison.OrdinalIgnoreCase)))
            {
                if (parentSet.Contains(drawing))
                {
                    continue;
                }

                AddRow(drawing, child: "", childCount: 0, exactLocations, fileNames, allRows, orphanRows);
                Log(log, $"2D orphan found from input file 1: {drawing}");
            }

            var outputPath = Path.Combine(outputFolder, $"BOM_Analysis_Output_{timestamp}.xlsx");
            ExcelWriter.Write(outputPath,
            [
                new ExcelSheet("BOM Analysis", Headers, allRows),
                new ExcelSheet("2D_Orphans", Headers, orphanRows),
                new ExcelSheet("2D_WithMultiple_3D", Headers, multiRows)
            ]);

            Log(log, $"Location entries read: {locations.Count}");
            Log(log, $"Relationship rows read: {relationRows.Count}");
            Log(log, $"Report rows generated: {allRows.Count}");
            Log(log, $"Rows without children in relationship file: {parentOnlyRows}");
            Log(log, $"2D orphan rows generated: {orphanRows.Count}");
            Log(log, $"Multiple 3D rows generated: {multiRows.Count}");
            Log(log, $"Excel output: {outputPath}");
            Log(log, "Completed successfully.");

            Console.WriteLine();
            Console.WriteLine("Completed successfully.");
            Console.WriteLine($"Excel output: {outputPath}");
            Console.WriteLine($"Log file:     {logPath}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine("Failed to generate report.");
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }

    private static void AddRow(
        string parent,
        string child,
        int childCount,
        HashSet<string> exactLocations,
        HashSet<string> fileNames,
        List<string[]> allRows,
        List<string[]> orphanRows)
    {
        var row = CreateReportRow(parent, child, childCount, exactLocations, fileNames);
        allRows.Add(row);
        orphanRows.Add(row);
    }

    private static string[] CreateReportRow(
        string parent,
        string child,
        int childCount,
        HashSet<string> exactLocations,
        HashSet<string> fileNames)
    {
        var hasChild = IsPresent(child);
        var childName = hasChild ? Path.GetFileName(child) : "";

        return
        [
            parent,
            Path.GetFileName(parent),
            Path.GetExtension(parent),
            ExtractRevision(parent),
            childCount == 0 ? "TRUE" : "FALSE",
            childCount > 1 ? "TRUE" : "FALSE",
            child,
            childName,
            hasChild ? Path.GetExtension(child) : "",
            hasChild ? ExtractRevision(child) : "",
            hasChild ? "1" : "",
            hasChild ? Bool(exactLocations.Contains(child)) : "",
            hasChild ? Bool(fileNames.Contains(childName)) : ""
        ];
    }

    private static string PromptForExistingFile(string prompt)
    {
        while (true)
        {
            Console.Write($"{prompt}: ");
            var value = (Console.ReadLine() ?? "").Trim().Trim('"');
            if (File.Exists(value))
            {
                return value;
            }

            Console.WriteLine("File not found. Please enter a valid file path.");
        }
    }

    private static string PromptForOutputFolder(string prompt)
    {
        while (true)
        {
            Console.Write($"{prompt}: ");
            var value = (Console.ReadLine() ?? "").Trim().Trim('"');
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            Console.WriteLine("Please enter a valid folder path.");
        }
    }

    private static string ExtractRevision(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        if (string.IsNullOrWhiteSpace(name))
        {
            return "";
        }

        foreach (var marker in new[] { "_REV", "-REV", " REV " })
        {
            var index = name.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
            {
                return name[(index + marker.Length)..].Trim(' ', '_', '-');
            }
        }

        return "";
    }

    private static void Log(TextWriter writer, string message)
    {
        writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {Clean(message)}");
        writer.Flush();
    }

    private static string Bool(bool value) => value ? "TRUE" : "FALSE";
    private static bool IsPresent(string? value) => !string.IsNullOrWhiteSpace(value);
    private static string Clean(string value) => new string(value.Where(XmlConvert.IsXmlChar).ToArray()).Trim().Trim('"');
}

internal static class InputReader
{
    public static IEnumerable<IReadOnlyList<string>> ReadRows(string path) =>
        Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? ReadXlsxRows(path)
            : ReadTextRows(path);

    private static IEnumerable<IReadOnlyList<string>> ReadTextRows(string path)
    {
        foreach (var line in File.ReadLines(path, Encoding.UTF8))
        {
            if (!string.IsNullOrWhiteSpace(line))
            {
                yield return line.Contains('\t') ? line.Split('\t', StringSplitOptions.TrimEntries) : [line];
            }
        }
    }

    private static IEnumerable<IReadOnlyList<string>> ReadXlsxRows(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var strings = SharedStrings(archive);
        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidOperationException("The XLSX file does not contain xl/worksheets/sheet1.xml.");

        using var stream = sheetEntry.Open();
        var sheet = XDocument.Load(stream);
        var ns = sheet.Root?.Name.Namespace ?? XNamespace.None;

        foreach (var row in sheet.Descendants(ns + "row"))
        {
            var cells = row.Elements(ns + "c").Select(cell => CellValue(cell, ns, strings)).ToList();
            if (cells.Any(static cell => !string.IsNullOrWhiteSpace(cell)))
            {
                yield return cells;
            }
        }
    }

    private static List<string> SharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return [];
        }

        using var stream = entry.Open();
        var document = XDocument.Load(stream);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;
        return document.Descendants(ns + "si")
            .Select(si => string.Concat(si.Descendants(ns + "t").Select(static t => t.Value)))
            .ToList();
    }

    private static string CellValue(XElement cell, XNamespace ns, IReadOnlyList<string> strings)
    {
        var type = cell.Attribute("t")?.Value;
        var value = cell.Element(ns + "v")?.Value ?? "";

        if (type == "inlineStr")
        {
            return string.Concat(cell.Descendants(ns + "t").Select(static t => t.Value));
        }

        return type == "s" && int.TryParse(value, out var index) && index >= 0 && index < strings.Count
            ? strings[index]
            : value;
    }
}

internal sealed record ExcelSheet(string Name, IReadOnlyList<string> Headers, IReadOnlyList<string[]> Rows);

internal static class ExcelWriter
{
    private static readonly XNamespace ContentNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace SheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace CoreNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DctermsNs = "http://purl.org/dc/terms/";
    private static readonly XNamespace XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

    public static void Write(string path, IReadOnlyList<ExcelSheet> sheets)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }

        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", ContentTypes(sheets.Count));
        Add(archive, "_rels/.rels", PackageRels());
        Add(archive, "docProps/app.xml", AppProps());
        Add(archive, "docProps/core.xml", CoreProps());
        Add(archive, "xl/workbook.xml", Workbook(sheets));
        Add(archive, "xl/_rels/workbook.xml.rels", WorkbookRels(sheets.Count));
        Add(archive, "xl/styles.xml", Styles());

        for (var i = 0; i < sheets.Count; i++)
        {
            Add(archive, $"xl/worksheets/sheet{i + 1}.xml", Worksheet(sheets[i]));
        }
    }

    private static void Add(ZipArchive archive, string path, XDocument document)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false });
        document.Save(writer);
    }

    private static XDocument ContentTypes(int sheetCount) => Doc(new XElement(ContentNs + "Types",
        El(ContentNs, "Default", ("Extension", "rels"), ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
        El(ContentNs, "Default", ("Extension", "xml"), ("ContentType", "application/xml")),
        Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
        Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        Override("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        Override("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        Enumerable.Range(1, sheetCount).Select(i => Override($"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))));

    private static XElement Override(string part, string type) => El(ContentNs, "Override", ("PartName", part), ("ContentType", type));

    private static XDocument PackageRels() => Rels(
        Rel("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml"),
        Rel("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
        Rel("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"));

    private static XDocument WorkbookRels(int sheetCount) => Rels(
        Enumerable.Range(1, sheetCount).Select(i => Rel($"rId{i}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $"worksheets/sheet{i}.xml"))
            .Append(Rel($"rId{sheetCount + 1}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml")).ToArray());

    private static XDocument Workbook(IReadOnlyList<ExcelSheet> sheets) => Doc(new XElement(SheetNs + "workbook",
        new XAttribute(XNamespace.Xmlns + "r", OfficeRelNs),
        new XElement(SheetNs + "bookViews", new XElement(SheetNs + "workbookView")),
        new XElement(SheetNs + "sheets", sheets.Select((s, i) => new XElement(SheetNs + "sheet",
            new XAttribute("name", SheetName(s.Name)),
            new XAttribute("sheetId", i + 1),
            new XAttribute(OfficeRelNs + "id", $"rId{i + 1}"))))));

    private static XDocument AppProps()
    {
        XNamespace app = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        return Doc(new XElement(app + "Properties",
            new XElement(app + "Application", "BOM Analysis"),
            new XElement(app + "DocSecurity", "0"),
            new XElement(app + "ScaleCrop", "false")));
    }

    private static XDocument CoreProps()
    {
        var now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        return Doc(new XElement(CoreNs + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "dc", DcNs),
            new XAttribute(XNamespace.Xmlns + "dcterms", DctermsNs),
            new XAttribute(XNamespace.Xmlns + "xsi", XsiNs),
            new XElement(DcNs + "creator", "BOM Analysis"),
            new XElement(CoreNs + "lastModifiedBy", "BOM Analysis"),
            new XElement(DctermsNs + "created", new XAttribute(XsiNs + "type", "dcterms:W3CDTF"), now),
            new XElement(DctermsNs + "modified", new XAttribute(XsiNs + "type", "dcterms:W3CDTF"), now)));
    }

    private static XDocument Styles() => Doc(new XElement(SheetNs + "styleSheet",
        new XElement(SheetNs + "fonts", new XAttribute("count", 2), Font(false), Font(true)),
        new XElement(SheetNs + "fills", new XAttribute("count", 2), Fill("none"), Fill("gray125")),
        new XElement(SheetNs + "borders", new XAttribute("count", 1),
            new XElement(SheetNs + "border", new XElement(SheetNs + "left"), new XElement(SheetNs + "right"), new XElement(SheetNs + "top"), new XElement(SheetNs + "bottom"), new XElement(SheetNs + "diagonal"))),
        new XElement(SheetNs + "cellStyleXfs", new XAttribute("count", 1), Xf(0)),
        new XElement(SheetNs + "cellXfs", new XAttribute("count", 2), Xf(0), Xf(1, true)),
        new XElement(SheetNs + "cellStyles", new XAttribute("count", 1), El(SheetNs, "cellStyle", ("name", "Normal"), ("xfId", "0"), ("builtinId", "0")))));

    private static XDocument Worksheet(ExcelSheet sheet)
    {
        var rows = new[] { sheet.Headers }.Concat(sheet.Rows).Select((row, index) => Row(index + 1, row));
        return Doc(new XElement(SheetNs + "worksheet",
            El(SheetNs, "dimension", ("ref", $"A1:M{sheet.Rows.Count + 1}")),
            new XElement(SheetNs + "cols", Col(1, 1, 60), Col(2, 3, 24), Col(4, 6, 14), Col(7, 7, 60), Col(8, 10, 24), Col(11, 13, 18)),
            new XElement(SheetNs + "sheetData", rows),
            El(SheetNs, "autoFilter", ("ref", $"A1:M{sheet.Rows.Count + 1}"))));
    }

    private static XElement Row(int rowNumber, IReadOnlyList<string> cells) =>
        new(SheetNs + "row", new XAttribute("r", rowNumber), cells.Select((value, i) => Cell(rowNumber, i + 1, value, rowNumber == 1)));

    private static XElement Cell(int row, int col, string value, bool header) =>
        new(SheetNs + "c",
            new XAttribute("r", $"{ColumnName(col)}{row}"),
            new XAttribute("t", "inlineStr"),
            header ? new XAttribute("s", 1) : null,
            new XElement(SheetNs + "is", new XElement(SheetNs + "t", Clean(value))));

    private static XDocument Rels(params XElement[] relationships) => Doc(new XElement(PackageRelNs + "Relationships", relationships));
    private static XElement Rel(string id, string type, string target) => El(PackageRelNs, "Relationship", ("Id", id), ("Type", type), ("Target", target));
    private static XElement Font(bool bold) => new(SheetNs + "font", bold ? new XElement(SheetNs + "b") : null, El(SheetNs, "sz", ("val", "11")), El(SheetNs, "name", ("val", "Calibri")));
    private static XElement Fill(string pattern) => new(SheetNs + "fill", El(SheetNs, "patternFill", ("patternType", pattern)));
    private static XElement Xf(int fontId, bool applyFont = false) => El(SheetNs, "xf", ("numFmtId", "0"), ("fontId", fontId.ToString(CultureInfo.InvariantCulture)), ("fillId", "0"), ("borderId", "0"), ("xfId", "0"), ("applyFont", applyFont ? "1" : ""));
    private static XElement Col(int min, int max, double width) => El(SheetNs, "col", ("min", $"{min}"), ("max", $"{max}"), ("width", width.ToString(CultureInfo.InvariantCulture)), ("customWidth", "1"));
    private static XDocument Doc(XElement root) => new(new XDeclaration("1.0", "utf-8", "yes"), root);
    private static XElement El(XNamespace ns, string name, params (string Name, string Value)[] attrs) => new(ns + name, attrs.Where(a => a.Value != "").Select(a => new XAttribute(a.Name, a.Value)));

    private static string SheetName(string name)
    {
        var invalid = new HashSet<char>([':', '\\', '/', '?', '*', '[', ']']);
        var clean = new string(Clean(name).Where(c => !invalid.Contains(c)).Take(31).ToArray());
        return string.IsNullOrWhiteSpace(clean) ? "Sheet" : clean;
    }

    private static string ColumnName(int column)
    {
        var name = "";
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }

        return name;
    }

    private static string Clean(string value) => new(value.Where(XmlConvert.IsXmlChar).ToArray());
}
