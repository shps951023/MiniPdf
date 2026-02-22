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
}
