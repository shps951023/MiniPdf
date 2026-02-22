using System.IO.Compression;
using System.Xml.Linq;

namespace MiniPdf;

/// <summary>
/// Reads basic text data from Excel (.xlsx) files.
/// Supports reading cell values (strings and numbers) without external dependencies.
/// </summary>
internal static class ExcelReader
{
    /// <summary>
    /// Reads all sheets from an Excel file and returns their data as a list of sheets,
    /// where each sheet is a list of rows, and each row is a list of cell values.
    /// </summary>
    internal static List<ExcelSheet> ReadSheets(Stream stream)
    {
        var sheets = new List<ExcelSheet>();

        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // Read shared strings table
        var sharedStrings = ReadSharedStrings(archive);

        // Read workbook to get sheet names and order
        var sheetInfos = ReadWorkbook(archive);

        // Read each sheet
        foreach (var info in sheetInfos)
        {
            var entry = archive.GetEntry($"xl/worksheets/sheet{info.SheetId}.xml")
                        ?? archive.GetEntry($"xl/worksheets/{info.Name}.xml");

            // Try by relationship id pattern
            entry ??= archive.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
                e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

            if (entry == null) continue;

            var rows = ReadSheet(entry, sharedStrings);
            sheets.Add(new ExcelSheet(info.Name, rows));
        }

        // If no sheets found via workbook, try reading sheet1 directly
        if (sheets.Count == 0)
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
            if (entry != null)
            {
                var rows = ReadSheet(entry, sharedStrings);
                sheets.Add(new ExcelSheet("Sheet1", rows));
            }
        }

        return sheets;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var strings = new List<string>();
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return strings;

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        foreach (var si in doc.Descendants(ns + "si"))
        {
            // Handle both simple <t> and rich text <r><t> patterns
            var text = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
            strings.Add(text);
        }

        return strings;
    }

    private static List<SheetInfo> ReadWorkbook(ZipArchive archive)
    {
        var result = new List<SheetInfo>();
        var entry = archive.GetEntry("xl/workbook.xml");
        if (entry == null) return result;

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        var sheetId = 1;
        foreach (var sheet in doc.Descendants(ns + "sheet"))
        {
            var name = sheet.Attribute("name")?.Value ?? $"Sheet{sheetId}";
            result.Add(new SheetInfo(name, sheetId));
            sheetId++;
        }

        return result;
    }

    private static List<List<string>> ReadSheet(ZipArchiveEntry entry, List<string> sharedStrings)
    {
        var rows = new List<List<string>>();

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        foreach (var row in doc.Descendants(ns + "row"))
        {
            var cells = new List<string>();
            var lastColIndex = 0;

            foreach (var cell in row.Elements(ns + "c"))
            {
                // Parse column reference to handle gaps (e.g., A1, C1 means B is empty)
                var reference = cell.Attribute("r")?.Value ?? "";
                var colIndex = ParseColumnIndex(reference);

                // Fill empty cells for gaps
                while (lastColIndex < colIndex)
                {
                    cells.Add(string.Empty);
                    lastColIndex++;
                }

                var type = cell.Attribute("t")?.Value;
                var value = cell.Element(ns + "v")?.Value ?? "";

                if (type == "s" && int.TryParse(value, out var idx) && idx < sharedStrings.Count)
                {
                    cells.Add(sharedStrings[idx]);
                }
                else if (type == "inlineStr")
                {
                    var inlineText = string.Concat(cell.Descendants(ns + "t").Select(t => t.Value));
                    cells.Add(inlineText);
                }
                else
                {
                    cells.Add(value);
                }

                lastColIndex = colIndex + 1;
            }

            rows.Add(cells);
        }

        return rows;
    }

    private static int ParseColumnIndex(string cellReference)
    {
        var col = 0;
        foreach (var c in cellReference)
        {
            if (char.IsLetter(c))
            {
                col = col * 26 + (char.ToUpper(c) - 'A' + 1);
            }
            else
            {
                break;
            }
        }
        return col > 0 ? col - 1 : 0;
    }

    internal record SheetInfo(string Name, int SheetId);
}

/// <summary>
/// Represents a sheet read from an Excel file.
/// </summary>
internal sealed class ExcelSheet
{
    public string Name { get; }
    public List<List<string>> Rows { get; }

    internal ExcelSheet(string name, List<List<string>> rows)
    {
        Name = name;
        Rows = rows;
    }
}
