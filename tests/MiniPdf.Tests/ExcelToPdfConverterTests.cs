using System.IO.Compression;
using System.Text;

namespace MiniPdf.Tests;

public class ExcelToPdfConverterTests
{
    [Fact]
    public void Convert_SimpleExcel_ProducesValidPdf()
    {
        using var excelStream = CreateSimpleExcel(new[]
        {
            new[] { "Name", "Age", "City" },
            new[] { "Alice", "30", "New York" },
            new[] { "Bob", "25", "London" },
        });

        var doc = ExcelToPdfConverter.Convert(excelStream);
        var bytes = doc.ToArray();
        var content = Encoding.ASCII.GetString(bytes);

        Assert.StartsWith("%PDF-1.4", content);
        Assert.Contains("Name", content);
        Assert.Contains("Alice", content);
        Assert.Contains("Bob", content);
        Assert.Contains("%%EOF", content);
    }

    [Fact]
    public void Convert_WithOptions_UsesCustomSettings()
    {
        using var excelStream = CreateSimpleExcel(new[]
        {
            new[] { "Header1", "Header2" },
            new[] { "Value1", "Value2" },
        });

        var options = new ExcelToPdfConverter.ConversionOptions
        {
            FontSize = 14,
            MarginLeft = 72,
            PageWidth = 595, // A4
            PageHeight = 842, // A4
            IncludeSheetName = false,
        };

        var doc = ExcelToPdfConverter.Convert(excelStream, options);
        Assert.True(doc.Pages.Count >= 1);
        var bytes = doc.ToArray();
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void Convert_EmptyExcel_CreatesAtLeastOnePage()
    {
        using var excelStream = CreateSimpleExcel(Array.Empty<string[]>());

        var doc = ExcelToPdfConverter.Convert(excelStream);
        Assert.True(doc.Pages.Count >= 1);
    }

    [Fact]
    public void ConvertToFile_CreatesOutputFile()
    {
        var excelPath = Path.Combine(Path.GetTempPath(), $"minipdf_test_{Guid.NewGuid()}.xlsx");
        var pdfPath = Path.Combine(Path.GetTempPath(), $"minipdf_test_{Guid.NewGuid()}.pdf");

        try
        {
            using (var fs = File.Create(excelPath))
            using (var excelStream = CreateSimpleExcel(new[]
            {
                new[] { "Test", "Data" },
                new[] { "1", "2" },
            }))
            {
                excelStream.CopyTo(fs);
            }

            ExcelToPdfConverter.ConvertToFile(excelPath, pdfPath);

            Assert.True(File.Exists(pdfPath));
            var bytes = File.ReadAllBytes(pdfPath);
            Assert.StartsWith("%PDF-1.4", Encoding.ASCII.GetString(bytes));
        }
        finally
        {
            if (File.Exists(excelPath)) File.Delete(excelPath);
            if (File.Exists(pdfPath)) File.Delete(pdfPath);
        }
    }

    [Fact]
    public void Convert_ManyRows_CreatesMultiplePages()
    {
        var rows = new List<string[]>();
        for (var i = 0; i < 100; i++)
        {
            rows.Add(new[] { $"Row{i}", $"Value{i}", $"Data{i}" });
        }

        using var excelStream = CreateSimpleExcel(rows.ToArray());
        var doc = ExcelToPdfConverter.Convert(excelStream);

        // 100 rows at ~14pt line height should require multiple pages
        Assert.True(doc.Pages.Count >= 2, $"Expected at least 2 pages, got {doc.Pages.Count}");
    }

    [Fact]
    public void Convert_WithTextColor_PreservesColorInPdf()
    {
        // Create an xlsx with red text in cell A1
        using var excelStream = CreateColoredExcel(
            new[] { ("Red Text", "FFFF0000"), ("Normal", "") },
            new[] { ("Blue Val", "FF0000FF"), ("Green", "FF00FF00") }
        );

        var doc = ExcelToPdfConverter.Convert(excelStream);
        var bytes = doc.ToArray();
        var content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("Red Text", content);
        Assert.Contains("Blue Val", content);
        // Verify that non-black color operators appear (red = 1.000 0.000 0.000 rg)
        Assert.Contains("1.000 0.000 0.000 rg", content);
        // Blue = 0.000 0.000 1.000 rg
        Assert.Contains("0.000 0.000 1.000 rg", content);
    }

    /// <summary>
    /// Creates a minimal valid .xlsx file in memory with the given data.
    /// </summary>
    private static MemoryStream CreateSimpleExcel(string[][] rows)
    {
        var ms = new MemoryStream();

        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            // [Content_Types].xml
            AddEntry(archive, "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                </Types>
                """);

            // _rels/.rels
            AddEntry(archive, "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            // xl/_rels/workbook.xml.rels
            AddEntry(archive, "xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                </Relationships>
                """);

            // xl/workbook.xml
            AddEntry(archive, "xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                  </sheets>
                </workbook>
                """);

            // Build shared strings and sheet data
            var sharedStrings = new List<string>();
            var sharedStringIndex = new Dictionary<string, int>();

            int GetStringIndex(string value)
            {
                if (!sharedStringIndex.TryGetValue(value, out var idx))
                {
                    idx = sharedStrings.Count;
                    sharedStrings.Add(value);
                    sharedStringIndex[value] = idx;
                }
                return idx;
            }

            // Build sheet XML
            var sheetSb = new StringBuilder();
            sheetSb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
            sheetSb.AppendLine("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
            sheetSb.AppendLine("<sheetData>");

            for (var r = 0; r < rows.Length; r++)
            {
                sheetSb.AppendLine($"  <row r=\"{r + 1}\">");
                for (var c = 0; c < rows[r].Length; c++)
                {
                    var colLetter = (char)('A' + c);
                    var cellRef = $"{colLetter}{r + 1}";
                    var idx = GetStringIndex(rows[r][c]);
                    sheetSb.AppendLine($"    <c r=\"{cellRef}\" t=\"s\"><v>{idx}</v></c>");
                }
                sheetSb.AppendLine("  </row>");
            }

            sheetSb.AppendLine("</sheetData>");
            sheetSb.AppendLine("</worksheet>");

            AddEntry(archive, "xl/worksheets/sheet1.xml", sheetSb.ToString());

            // Build shared strings XML
            var ssSb = new StringBuilder();
            ssSb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
            ssSb.AppendLine($"""<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{sharedStrings.Count}" uniqueCount="{sharedStrings.Count}">""");
            foreach (var s in sharedStrings)
            {
                ssSb.AppendLine($"  <si><t>{EscapeXml(s)}</t></si>");
            }
            ssSb.AppendLine("</sst>");

            AddEntry(archive, "xl/sharedStrings.xml", ssSb.ToString());
        }

        ms.Position = 0;
        return ms;
    }

    private static void AddEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
        writer.Write(content);
    }

    private static string EscapeXml(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }

    /// <summary>
    /// Creates a minimal .xlsx with per-cell font colors.
    /// Each row is an array of (text, argbHex) tuples. Empty argb = default/black.
    /// </summary>
    private static MemoryStream CreateColoredExcel(params (string text, string argb)[][] rows)
    {
        var ms = new MemoryStream();

        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(archive, "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """);

            AddEntry(archive, "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

            AddEntry(archive, "xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);

            AddEntry(archive, "xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                  </sheets>
                </workbook>
                """);

            // Collect unique colors and build fonts + cellXfs
            var colorToFontIndex = new Dictionary<string, int>();
            var fontsSb = new StringBuilder();
            var xfsSb = new StringBuilder();

            // Font 0 / xf 0 = default (black, no color element)
            fontsSb.AppendLine("  <font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
            xfsSb.AppendLine("  <xf fontId=\"0\"/>");
            var nextFontId = 1;
            var nextXfId = 1;

            // Map (fontIndex -> xfIndex) for non-default colors
            var colorToXfIndex = new Dictionary<string, int>();

            foreach (var row in rows)
            {
                foreach (var (_, argb) in row)
                {
                    if (string.IsNullOrEmpty(argb) || colorToXfIndex.ContainsKey(argb))
                        continue;

                    fontsSb.AppendLine($"  <font><color rgb=\"{argb}\"/><sz val=\"11\"/><name val=\"Calibri\"/></font>");
                    colorToFontIndex[argb] = nextFontId;

                    xfsSb.AppendLine($"  <xf fontId=\"{nextFontId}\"/>");
                    colorToXfIndex[argb] = nextXfId;

                    nextFontId++;
                    nextXfId++;
                }
            }

            var stylesSb = new StringBuilder();
            stylesSb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
            stylesSb.AppendLine("""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
            stylesSb.AppendLine($"<fonts count=\"{nextFontId}\">");
            stylesSb.Append(fontsSb);
            stylesSb.AppendLine("</fonts>");
            stylesSb.AppendLine($"<cellXfs count=\"{nextXfId}\">");
            stylesSb.Append(xfsSb);
            stylesSb.AppendLine("</cellXfs>");
            stylesSb.AppendLine("</styleSheet>");

            AddEntry(archive, "xl/styles.xml", stylesSb.ToString());

            // Shared strings
            var sharedStrings = new List<string>();
            var sharedStringIndex = new Dictionary<string, int>();

            int GetStringIndex(string value)
            {
                if (!sharedStringIndex.TryGetValue(value, out var idx))
                {
                    idx = sharedStrings.Count;
                    sharedStrings.Add(value);
                    sharedStringIndex[value] = idx;
                }
                return idx;
            }

            // Sheet data
            var sheetSb = new StringBuilder();
            sheetSb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
            sheetSb.AppendLine("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
            sheetSb.AppendLine("<sheetData>");

            for (var r = 0; r < rows.Length; r++)
            {
                sheetSb.AppendLine($"  <row r=\"{r + 1}\">");
                for (var c = 0; c < rows[r].Length; c++)
                {
                    var colLetter = (char)('A' + c);
                    var cellRef = $"{colLetter}{r + 1}";
                    var idx = GetStringIndex(rows[r][c].text);
                    var styleIdx = !string.IsNullOrEmpty(rows[r][c].argb) && colorToXfIndex.TryGetValue(rows[r][c].argb, out var si) ? si : 0;
                    sheetSb.AppendLine($"    <c r=\"{cellRef}\" t=\"s\" s=\"{styleIdx}\"><v>{idx}</v></c>");
                }
                sheetSb.AppendLine("  </row>");
            }

            sheetSb.AppendLine("</sheetData>");
            sheetSb.AppendLine("</worksheet>");

            AddEntry(archive, "xl/worksheets/sheet1.xml", sheetSb.ToString());

            var ssSb = new StringBuilder();
            ssSb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
            ssSb.AppendLine($"""<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{sharedStrings.Count}" uniqueCount="{sharedStrings.Count}">""");
            foreach (var s in sharedStrings)
            {
                ssSb.AppendLine($"  <si><t>{EscapeXml(s)}</t></si>");
            }
            ssSb.AppendLine("</sst>");

            AddEntry(archive, "xl/sharedStrings.xml", ssSb.ToString());
        }

        ms.Position = 0;
        return ms;
    }
}
