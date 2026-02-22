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

        // Read styles (font colors)
        var fontColors = ReadFontColors(archive);
        var cellXfFontIndices = ReadCellXfFontIndices(archive);

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

            var rows = ReadSheet(entry, sharedStrings, fontColors, cellXfFontIndices);
            sheets.Add(new ExcelSheet(info.Name, rows));
        }

        // If no sheets found via workbook, try reading sheet1 directly
        if (sheets.Count == 0)
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
            if (entry != null)
            {
                var rows = ReadSheet(entry, sharedStrings, fontColors, cellXfFontIndices);
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

    private static List<PdfColor?> ReadFontColors(ZipArchive archive)
    {
        var colors = new List<PdfColor?>();
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null) return colors;

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        // Read <fonts> -> <font> elements
        var fontsElement = doc.Descendants(ns + "fonts").FirstOrDefault();
        if (fontsElement == null) return colors;

        foreach (var font in fontsElement.Elements(ns + "font"))
        {
            var colorEl = font.Element(ns + "color");
            if (colorEl == null)
            {
                colors.Add(null);
                continue;
            }

            // Try rgb attribute (ARGB hex, e.g., "FF0000FF")
            var rgb = colorEl.Attribute("rgb")?.Value;
            if (!string.IsNullOrEmpty(rgb))
            {
                colors.Add(PdfColor.FromHex(rgb));
                continue;
            }

            // Try theme attribute (would need theme parsing - skip for now)
            // Try indexed attribute
            var indexed = colorEl.Attribute("indexed")?.Value;
            if (!string.IsNullOrEmpty(indexed) && int.TryParse(indexed, out var idx))
            {
                colors.Add(GetIndexedColor(idx));
                continue;
            }

            colors.Add(null);
        }

        return colors;
    }

    private static List<int> ReadCellXfFontIndices(ZipArchive archive)
    {
        var indices = new List<int>();
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null) return indices;

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        // Read <cellXfs> -> <xf> elements to map style index -> font index
        var cellXfs = doc.Descendants(ns + "cellXfs").FirstOrDefault();
        if (cellXfs == null) return indices;

        foreach (var xf in cellXfs.Elements(ns + "xf"))
        {
            var fontId = xf.Attribute("fontId")?.Value;
            indices.Add(int.TryParse(fontId, out var fid) ? fid : 0);
        }

        return indices;
    }

    private static PdfColor? GetIndexedColor(int index)
    {
        // Standard Excel indexed colors (subset of the 64 built-in colors)
        return index switch
        {
            0 => PdfColor.FromRgb(0, 0, 0),        // Black
            1 => PdfColor.FromRgb(255, 255, 255),   // White
            2 => PdfColor.FromRgb(255, 0, 0),       // Red
            3 => PdfColor.FromRgb(0, 255, 0),       // Green
            4 => PdfColor.FromRgb(0, 0, 255),       // Blue
            5 => PdfColor.FromRgb(255, 255, 0),     // Yellow
            6 => PdfColor.FromRgb(255, 0, 255),     // Magenta
            7 => PdfColor.FromRgb(0, 255, 255),     // Cyan
            8 => PdfColor.FromRgb(0, 0, 0),         // Black
            9 => PdfColor.FromRgb(255, 255, 255),   // White
            10 => PdfColor.FromRgb(255, 0, 0),      // Red
            11 => PdfColor.FromRgb(0, 255, 0),      // Green
            12 => PdfColor.FromRgb(0, 0, 255),      // Blue
            13 => PdfColor.FromRgb(255, 255, 0),    // Yellow
            14 => PdfColor.FromRgb(255, 0, 255),    // Magenta
            15 => PdfColor.FromRgb(0, 255, 255),    // Cyan
            16 => PdfColor.FromRgb(128, 0, 0),      // Dark Red
            17 => PdfColor.FromRgb(0, 128, 0),      // Dark Green
            18 => PdfColor.FromRgb(0, 0, 128),      // Dark Blue
            19 => PdfColor.FromRgb(128, 128, 0),    // Olive
            20 => PdfColor.FromRgb(128, 0, 128),    // Purple
            21 => PdfColor.FromRgb(0, 128, 128),    // Teal
            22 => PdfColor.FromRgb(192, 192, 192),  // Silver
            23 => PdfColor.FromRgb(128, 128, 128),  // Grey
            _ => null
        };
    }

    private static PdfColor? ResolveCellColor(int styleIndex, List<PdfColor?> fontColors, List<int> cellXfFontIndices)
    {
        if (styleIndex < 0 || styleIndex >= cellXfFontIndices.Count)
            return null;

        var fontIndex = cellXfFontIndices[styleIndex];
        if (fontIndex < 0 || fontIndex >= fontColors.Count)
            return null;

        return fontColors[fontIndex];
    }

    private static List<List<ExcelCell>> ReadSheet(ZipArchiveEntry entry, List<string> sharedStrings, List<PdfColor?> fontColors, List<int> cellXfFontIndices)
    {
        var rows = new List<List<ExcelCell>>();

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

        foreach (var row in doc.Descendants(ns + "row"))
        {
            var cells = new List<ExcelCell>();
            var lastColIndex = 0;

            foreach (var cell in row.Elements(ns + "c"))
            {
                // Parse column reference to handle gaps (e.g., A1, C1 means B is empty)
                var reference = cell.Attribute("r")?.Value ?? "";
                var colIndex = ParseColumnIndex(reference);

                // Fill empty cells for gaps
                while (lastColIndex < colIndex)
                {
                    cells.Add(new ExcelCell(string.Empty, null));
                    lastColIndex++;
                }

                var type = cell.Attribute("t")?.Value;
                var value = cell.Element(ns + "v")?.Value ?? "";

                // Resolve color from style index
                var styleAttr = cell.Attribute("s")?.Value;
                PdfColor? color = null;
                if (int.TryParse(styleAttr, out var styleIndex))
                {
                    color = ResolveCellColor(styleIndex, fontColors, cellXfFontIndices);
                }

                string text;
                if (type == "s" && int.TryParse(value, out var idx) && idx < sharedStrings.Count)
                {
                    text = sharedStrings[idx];
                }
                else if (type == "inlineStr")
                {
                    text = string.Concat(cell.Descendants(ns + "t").Select(t => t.Value));
                }
                else
                {
                    text = value;
                }

                cells.Add(new ExcelCell(text, color));
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
/// Represents a cell read from an Excel file.
/// </summary>
internal sealed record ExcelCell(string Text, PdfColor? Color);

/// <summary>
/// Represents a sheet read from an Excel file.
/// </summary>
internal sealed class ExcelSheet
{
    public string Name { get; }
    public List<List<ExcelCell>> Rows { get; }

    internal ExcelSheet(string name, List<List<ExcelCell>> rows)
    {
        Name = name;
        Rows = rows;
    }
}
