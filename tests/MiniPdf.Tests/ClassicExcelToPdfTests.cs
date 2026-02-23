using System.IO.Compression;
using System.Text;

namespace MiniPdf.Tests;

/// <summary>
/// 30 classic .xlsx conversion tests for Excel-to-PDF (issue #43).
/// Each test builds a minimal .xlsx in memory, converts it, and
/// asserts the output is a valid PDF containing the expected data.
/// </summary>
public class ClassicExcelToPdfTests
{
    // ── 1. Basic table with headers ────────────────────────────────────
    [Fact]
    public void Classic01_BasicTableWithHeaders()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Name", "Age", "City" },
            new[] { "Alice", "30", "New York" },
            new[] { "Bob", "25", "London" });

        AssertValidPdf(xlsx, "Name", "Alice", "Bob");
    }

    // ── 2. Multiple worksheets ─────────────────────────────────────────
    [Fact]
    public void Classic02_MultipleWorksheets()
    {
        using var xlsx = XlsxBuilder.MultiSheet(
            ("Sales", new[] { new[] { "Q1", "100" }, new[] { "Q2", "200" } }),
            ("Costs", new[] { new[] { "Rent", "500" }, new[] { "Salary", "3000" } }));

        var doc = ExcelToPdfConverter.Convert(xlsx);
        var pdf = PdfString(doc);

        Assert.StartsWith("%PDF-1.4", pdf);
        Assert.Contains("Q1", pdf);
        Assert.Contains("Rent", pdf);
        Assert.True(doc.Pages.Count >= 2);
    }

    // ── 3. Empty workbook (no data rows) ───────────────────────────────
    [Fact]
    public void Classic03_EmptyWorkbook()
    {
        using var xlsx = XlsxBuilder.Simple(Array.Empty<string[]>());
        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 1);
    }

    // ── 4. Single cell ─────────────────────────────────────────────────
    [Fact]
    public void Classic04_SingleCell()
    {
        using var xlsx = XlsxBuilder.Simple(new[] { "Hello" });
        AssertValidPdf(xlsx, "Hello");
    }

    // ── 5. Wide table (26 columns A–Z) ─────────────────────────────────
    [Fact]
    public void Classic05_WideTable()
    {
        var header = Enumerable.Range(0, 26)
                               .Select(i => ((char)('A' + i)).ToString())
                               .ToArray();
        var row = Enumerable.Range(1, 26).Select(i => i.ToString()).ToArray();

        using var xlsx = XlsxBuilder.Simple(header, row);
        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 1);
        Assert.True(doc.ToArray().Length > 0);
    }

    // ── 6. Tall table (200 rows → multi-page) ─────────────────────────
    [Fact]
    public void Classic06_TallTable()
    {
        var rows = Enumerable.Range(1, 200)
                             .Select(i => new[] { $"Row{i}", $"Val{i}" })
                             .ToArray();

        using var xlsx = XlsxBuilder.Simple(rows);
        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 3, $"Expected ≥3 pages, got {doc.Pages.Count}");
    }

    // ── 7. Numbers only ────────────────────────────────────────────────
    [Fact]
    public void Classic07_NumbersOnly()
    {
        using var xlsx = XlsxBuilder.SimpleNumbers(
            new[] { 1.0, 2.0, 3.0 },
            new[] { 4.0, 5.0, 6.0 });

        AssertValidPdf(xlsx);
    }

    // ── 8. Mixed text and numbers ──────────────────────────────────────
    [Fact]
    public void Classic08_MixedTextAndNumbers()
    {
        // text in shared strings, numbers as raw values
        using var xlsx = XlsxBuilder.MixedTextNumber(
            ("Item", 10.5),
            ("Tax", 0.08),
            ("Total", 10.58));

        AssertValidPdf(xlsx, "Item", "Tax", "Total");
    }

    // ── 9. Long text content (triggers truncation) ────────────────────
    [Fact]
    public void Classic09_LongText()
    {
        var longText = new string('X', 500);
        using var xlsx = XlsxBuilder.Simple(new[] { longText });
        AssertValidPdf(xlsx);
    }

    // ── 10. Special XML characters ─────────────────────────────────────
    [Fact]
    public void Classic10_SpecialXmlCharacters()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "A&B", "<tag>", "\"quoted\"", "it's" });
        AssertValidPdf(xlsx);
    }

    // ── 11. Sparse rows (gaps between data rows) ──────────────────────
    [Fact]
    public void Classic11_SparseRows()
    {
        // Rows 1, 5, 10 have data; others are absent from XML
        using var xlsx = XlsxBuilder.SparseRows(
            (1, new[] { "First" }),
            (5, new[] { "Fifth" }),
            (10, new[] { "Tenth" }));

        AssertValidPdf(xlsx, "First", "Fifth", "Tenth");
    }

    // ── 12. Sparse columns (A, D filled; B, C empty) ──────────────────
    [Fact]
    public void Classic12_SparseColumns()
    {
        // Cell A1 and D1 filled; B1 and C1 absent
        using var xlsx = XlsxBuilder.SparseColumns(
            (1, new (string col, string val)[] { ("A", "Left"), ("D", "Right") }));

        AssertValidPdf(xlsx, "Left", "Right");
    }

    // ── 13. Date-like strings ──────────────────────────────────────────
    [Fact]
    public void Classic13_DateStrings()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Date", "Event" },
            new[] { "2025-01-15", "Launch" },
            new[] { "2025-06-30", "Release" });

        AssertValidPdf(xlsx, "2025-01-15", "Launch");
    }

    // ── 14. Decimal numbers ────────────────────────────────────────────
    [Fact]
    public void Classic14_DecimalNumbers()
    {
        using var xlsx = XlsxBuilder.SimpleNumbers(
            new[] { 3.14159, 2.71828, 1.41421 });

        AssertValidPdf(xlsx);
    }

    // ── 15. Negative numbers ───────────────────────────────────────────
    [Fact]
    public void Classic15_NegativeNumbers()
    {
        using var xlsx = XlsxBuilder.SimpleNumbers(
            new[] { -100.0, -0.5, 0.0, 50.0 });

        AssertValidPdf(xlsx);
    }

    // ── 16. Percentage-like strings ────────────────────────────────────
    [Fact]
    public void Classic16_PercentageStrings()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Metric", "Rate" },
            new[] { "Conversion", "12.5%" },
            new[] { "Bounce", "45.3%" });

        AssertValidPdf(xlsx, "12.5%", "45.3%");
    }

    // ── 17. Currency-like strings ──────────────────────────────────────
    [Fact]
    public void Classic17_CurrencyStrings()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Item", "Price" },
            new[] { "Widget", "$19.99" },
            new[] { "Gadget", "$149.00" });

        AssertValidPdf(xlsx, "$19.99", "$149.00");
    }

    // ── 18. Stress test (1 000 rows × 10 cols) ────────────────────────
    [Fact]
    public void Classic18_LargeDataset()
    {
        var rows = Enumerable.Range(0, 1000)
                             .Select(r => Enumerable.Range(0, 10)
                                                    .Select(c => $"R{r}C{c}")
                                                    .ToArray())
                             .ToArray();

        using var xlsx = XlsxBuilder.Simple(rows);
        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 10, $"Expected ≥10 pages, got {doc.Pages.Count}");
    }

    // ── 19. Single column list ─────────────────────────────────────────
    [Fact]
    public void Classic19_SingleColumnList()
    {
        var rows = Enumerable.Range(1, 20)
                             .Select(i => new[] { $"Item {i}" })
                             .ToArray();

        using var xlsx = XlsxBuilder.Simple(rows);
        AssertValidPdf(xlsx, "Item 1", "Item 20");
    }

    // ── 20. Rows with all empty values ─────────────────────────────────
    [Fact]
    public void Classic20_AllEmptyCells()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "", "", "" },
            new[] { "", "", "" });

        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 1);
    }

    // ── 21. Header only (no data rows) ─────────────────────────────────
    [Fact]
    public void Classic21_HeaderOnly()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Col1", "Col2", "Col3" });

        AssertValidPdf(xlsx, "Col1", "Col2", "Col3");
    }

    // ── 22. Very long sheet name ───────────────────────────────────────
    [Fact]
    public void Classic22_LongSheetName()
    {
        var sheetName = "VeryLongSheetNameThatExceedsTypicalLength";
        using var xlsx = XlsxBuilder.MultiSheet(
            (sheetName, new[] { new[] { "Data" } }));

        var doc = ExcelToPdfConverter.Convert(xlsx);
        var pdf = PdfString(doc);
        Assert.Contains("Data", pdf);
    }

    // ── 23. Unicode / CJK text ─────────────────────────────────────────
    [Fact]
    public void Classic23_UnicodeText()
    {
        // Helvetica cannot render CJK, but the converter must not throw
        using var xlsx = XlsxBuilder.Simple(
            new[] { "English", "Chinese", "Japanese" },
            new[] { "Hello", "\u4F60\u597D", "\u3053\u3093\u306B\u3061\u306F" });

        var doc = ExcelToPdfConverter.Convert(xlsx);
        Assert.True(doc.Pages.Count >= 1);
    }

    // ── 24. Single color (red text) ────────────────────────────────────
    [Fact]
    public void Classic24_RedText()
    {
        using var xlsx = XlsxBuilder.Colored(
            new (string, string)[] { ("Error", "FFFF0000"), ("OK", "") });

        var pdf = PdfString(ExcelToPdfConverter.Convert(xlsx));
        Assert.Contains("Error", pdf);
        Assert.Contains("1.000 0.000 0.000 rg", pdf);   // red
    }

    // ── 25. Multiple colors ────────────────────────────────────────────
    [Fact]
    public void Classic25_MultipleColors()
    {
        using var xlsx = XlsxBuilder.Colored(
            new (string, string)[] { ("Red", "FFFF0000"), ("Green", "FF00FF00"), ("Blue", "FF0000FF") });

        var pdf = PdfString(ExcelToPdfConverter.Convert(xlsx));
        Assert.Contains("1.000 0.000 0.000 rg", pdf);
        Assert.Contains("0.000 0.000 1.000 rg", pdf);
    }

    // ── 26. Inline strings ─────────────────────────────────────────────
    [Fact]
    public void Classic26_InlineStrings()
    {
        using var xlsx = XlsxBuilder.InlineStrings("Inline1", "Inline2", "Inline3");
        AssertValidPdf(xlsx, "Inline1", "Inline2");
    }

    // ── 27. Single row (horizontal data) ───────────────────────────────
    [Fact]
    public void Classic27_SingleRow()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" });

        AssertValidPdf(xlsx, "Mon", "Sun");
    }

    // ── 28. Duplicate values ───────────────────────────────────────────
    [Fact]
    public void Classic28_DuplicateValues()
    {
        using var xlsx = XlsxBuilder.Simple(
            new[] { "Yes", "No", "Yes", "No" },
            new[] { "No", "Yes", "No", "Yes" });

        AssertValidPdf(xlsx, "Yes", "No");
    }

    // ── 29. Formula-result values (stored as plain numbers) ───────────
    [Fact]
    public void Classic29_FormulaResults()
    {
        // Formulas are stored as cached values in the <v> elements
        using var xlsx = XlsxBuilder.SimpleNumbers(
            new[] { 10.0, 20.0, 30.0 });   // =SUM(A1:B1) would cache 30

        AssertValidPdf(xlsx);
    }

    // ── 30. Mixed empty and filled sheets ──────────────────────────────
    [Fact]
    public void Classic30_MixedEmptyAndFilledSheets()
    {
        using var xlsx = XlsxBuilder.MultiSheet(
            ("Empty", Array.Empty<string[]>()),
            ("Data", new[] { new[] { "Hello", "World" } }),
            ("AlsoEmpty", Array.Empty<string[]>()));

        var doc = ExcelToPdfConverter.Convert(xlsx);
        var pdf = PdfString(doc);
        Assert.Contains("Hello", pdf);
        Assert.True(doc.Pages.Count >= 1);
    }

    // ────────────────────────────────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────────────────────────────────

    private static string PdfString(PdfDocument doc)
        => Encoding.ASCII.GetString(doc.ToArray());

    private static void AssertValidPdf(Stream xlsx, params string[] expectedTexts)
    {
        var doc = ExcelToPdfConverter.Convert(xlsx);
        var pdf = PdfString(doc);

        Assert.StartsWith("%PDF-1.4", pdf);
        Assert.Contains("%%EOF", pdf);
        Assert.True(doc.Pages.Count >= 1);

        foreach (var text in expectedTexts)
            Assert.Contains(text, pdf);
    }

    // ────────────────────────────────────────────────────────────────────
    // XlsxBuilder – tiny in-memory .xlsx factory
    // ────────────────────────────────────────────────────────────────────

    private static class XlsxBuilder
    {
        // ── Simple: shared-string rows on a single sheet ───────────────
        public static MemoryStream Simple(params string[][] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false);
                WriteWorkbook(z, "Sheet1");
                var (sheetXml, ssXml) = BuildSheetAndStrings(rows);
                AddEntry(z, "xl/worksheets/sheet1.xml", sheetXml);
                AddEntry(z, "xl/sharedStrings.xml", ssXml);
            }
            ms.Position = 0;
            return ms;
        }

        // ── SimpleNumbers: numeric cells (type omitted → number) ──────
        public static MemoryStream SimpleNumbers(params double[][] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z, hasSharedStrings: false);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false, hasSharedStrings: false);
                WriteWorkbook(z, "Sheet1");

                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                for (var r = 0; r < rows.Length; r++)
                {
                    sb.AppendLine($"  <row r=\"{r + 1}\">");
                    for (var c = 0; c < rows[r].Length; c++)
                    {
                        var col = ColLetter(c);
                        sb.AppendLine($"    <c r=\"{col}{r + 1}\"><v>{rows[r][c].ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                    }
                    sb.AppendLine("  </row>");
                }
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
            }
            ms.Position = 0;
            return ms;
        }

        // ── MixedTextNumber: col-A = text, col-B = number ─────────────
        public static MemoryStream MixedTextNumber(params (string text, double number)[] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false);
                WriteWorkbook(z, "Sheet1");

                var ss = new SharedStrings();
                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                for (var r = 0; r < rows.Length; r++)
                {
                    var idx = ss.Add(rows[r].text);
                    sb.AppendLine($"  <row r=\"{r + 1}\">");
                    sb.AppendLine($"    <c r=\"A{r + 1}\" t=\"s\"><v>{idx}</v></c>");
                    sb.AppendLine($"    <c r=\"B{r + 1}\"><v>{rows[r].number.ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                    sb.AppendLine("  </row>");
                }
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
                AddEntry(z, "xl/sharedStrings.xml", ss.ToXml());
            }
            ms.Position = 0;
            return ms;
        }

        // ── SparseRows: rows at arbitrary row numbers ───────────────────
        public static MemoryStream SparseRows(params (int rowNum, string[] values)[] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false);
                WriteWorkbook(z, "Sheet1");

                var ss = new SharedStrings();
                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                foreach (var (rowNum, values) in rows)
                {
                    sb.AppendLine($"  <row r=\"{rowNum}\">");
                    for (var c = 0; c < values.Length; c++)
                    {
                        var idx = ss.Add(values[c]);
                        sb.AppendLine($"    <c r=\"{ColLetter(c)}{rowNum}\" t=\"s\"><v>{idx}</v></c>");
                    }
                    sb.AppendLine("  </row>");
                }
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
                AddEntry(z, "xl/sharedStrings.xml", ss.ToXml());
            }
            ms.Position = 0;
            return ms;
        }

        // ── SparseColumns: cells at arbitrary column letters ────────────
        public static MemoryStream SparseColumns(params (int rowNum, (string col, string val)[] cells)[] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false);
                WriteWorkbook(z, "Sheet1");

                var ss = new SharedStrings();
                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                foreach (var (rowNum, cells) in rows)
                {
                    sb.AppendLine($"  <row r=\"{rowNum}\">");
                    foreach (var (col, val) in cells)
                    {
                        var idx = ss.Add(val);
                        sb.AppendLine($"    <c r=\"{col}{rowNum}\" t=\"s\"><v>{idx}</v></c>");
                    }
                    sb.AppendLine("  </row>");
                }
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
                AddEntry(z, "xl/sharedStrings.xml", ss.ToXml());
            }
            ms.Position = 0;
            return ms;
        }

        // ── MultiSheet: multiple named sheets ──────────────────────────
        public static MemoryStream MultiSheet(params (string name, string[][] rows)[] sheets)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z, sheetCount: sheets.Length);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: sheets.Length, hasStyles: false);
                WriteWorkbook(z, sheets.Select(s => s.name).ToArray());

                // Each sheet gets its own shared strings scope to keep things simple.
                // We merge all into one global shared strings table.
                var globalSs = new SharedStrings();
                for (var i = 0; i < sheets.Length; i++)
                {
                    var (_, rows) = sheets[i];
                    var sb = new StringBuilder();
                    sb.AppendLine(XmlHeader);
                    sb.AppendLine(WsOpen);
                    sb.AppendLine("<sheetData>");
                    for (var r = 0; r < rows.Length; r++)
                    {
                        sb.AppendLine($"  <row r=\"{r + 1}\">");
                        for (var c = 0; c < rows[r].Length; c++)
                        {
                            var idx = globalSs.Add(rows[r][c]);
                            sb.AppendLine($"    <c r=\"{ColLetter(c)}{r + 1}\" t=\"s\"><v>{idx}</v></c>");
                        }
                        sb.AppendLine("  </row>");
                    }
                    sb.AppendLine("</sheetData></worksheet>");
                    AddEntry(z, $"xl/worksheets/sheet{i + 1}.xml", sb.ToString());
                }
                AddEntry(z, "xl/sharedStrings.xml", globalSs.ToXml());
            }
            ms.Position = 0;
            return ms;
        }

        // ── Colored: single sheet, each cell has optional ARGB colour ──
        public static MemoryStream Colored(params (string text, string argb)[][] rows)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z, hasStyles: true);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: true);
                WriteWorkbook(z, "Sheet1");

                // Build unique colour → font/xf mapping
                var colorMap = new Dictionary<string, int>();
                var fontsSb = new StringBuilder();
                var xfsSb = new StringBuilder();
                fontsSb.AppendLine("  <font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
                xfsSb.AppendLine("  <xf fontId=\"0\"/>");
                var nextId = 1;

                foreach (var row in rows)
                    foreach (var (_, argb) in row)
                        if (!string.IsNullOrEmpty(argb) && !colorMap.ContainsKey(argb))
                        {
                            fontsSb.AppendLine($"  <font><color rgb=\"{argb}\"/><sz val=\"11\"/><name val=\"Calibri\"/></font>");
                            xfsSb.AppendLine($"  <xf fontId=\"{nextId}\"/>");
                            colorMap[argb] = nextId++;
                        }

                var stySb = new StringBuilder();
                stySb.AppendLine(XmlHeader);
                stySb.AppendLine("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                stySb.AppendLine($"<fonts count=\"{nextId}\">");
                stySb.Append(fontsSb);
                stySb.AppendLine("</fonts>");
                stySb.AppendLine($"<cellXfs count=\"{nextId}\">");
                stySb.Append(xfsSb);
                stySb.AppendLine("</cellXfs>");
                stySb.AppendLine("</styleSheet>");
                AddEntry(z, "xl/styles.xml", stySb.ToString());

                var ss = new SharedStrings();
                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                for (var r = 0; r < rows.Length; r++)
                {
                    sb.AppendLine($"  <row r=\"{r + 1}\">");
                    for (var c = 0; c < rows[r].Length; c++)
                    {
                        var (text, argb) = rows[r][c];
                        var idx = ss.Add(text);
                        var sty = !string.IsNullOrEmpty(argb) && colorMap.TryGetValue(argb, out var si) ? si : 0;
                        sb.AppendLine($"    <c r=\"{ColLetter(c)}{r + 1}\" t=\"s\" s=\"{sty}\"><v>{idx}</v></c>");
                    }
                    sb.AppendLine("  </row>");
                }
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
                AddEntry(z, "xl/sharedStrings.xml", ss.ToXml());
            }
            ms.Position = 0;
            return ms;
        }

        // ── InlineStrings: uses inlineStr type instead of shared strings
        public static MemoryStream InlineStrings(params string[] values)
        {
            var ms = new MemoryStream();
            using (var z = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteContentTypes(z, hasSharedStrings: false);
                WriteRootRels(z);
                WriteWorkbookRels(z, sheetCount: 1, hasStyles: false, hasSharedStrings: false);
                WriteWorkbook(z, "Sheet1");

                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine(WsOpen);
                sb.AppendLine("<sheetData>");
                sb.AppendLine("  <row r=\"1\">");
                for (var c = 0; c < values.Length; c++)
                {
                    sb.AppendLine($"    <c r=\"{ColLetter(c)}1\" t=\"inlineStr\"><is><t>{Esc(values[c])}</t></is></c>");
                }
                sb.AppendLine("  </row>");
                sb.AppendLine("</sheetData></worksheet>");
                AddEntry(z, "xl/worksheets/sheet1.xml", sb.ToString());
            }
            ms.Position = 0;
            return ms;
        }

        // ── Internals ──────────────────────────────────────────────────

        private const string XmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
        private const string WsOpen = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

        private static string ColLetter(int index)
        {
            var result = "";
            var i = index;
            do
            {
                result = (char)('A' + i % 26) + result;
                i = i / 26 - 1;
            } while (i >= 0);
            return result;
        }

        private static string Esc(string s) => s
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");

        private static void AddEntry(ZipArchive z, string path, string content)
        {
            var e = z.CreateEntry(path);
            using var w = new StreamWriter(e.Open(), Encoding.UTF8);
            w.Write(content);
        }

        private static void WriteContentTypes(ZipArchive z, int sheetCount = 1, bool hasSharedStrings = true, bool hasStyles = false)
        {
            var sb = new StringBuilder();
            sb.AppendLine(XmlHeader);
            sb.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.AppendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.AppendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.AppendLine("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (var i = 1; i <= sheetCount; i++)
                sb.AppendLine($"  <Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            if (hasSharedStrings)
                sb.AppendLine("  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            if (hasStyles)
                sb.AppendLine("  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            sb.AppendLine("</Types>");
            AddEntry(z, "[Content_Types].xml", sb.ToString());
        }

        private static void WriteRootRels(ZipArchive z)
        {
            AddEntry(z, "_rels/.rels",
                XmlHeader + "\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
                "</Relationships>");
        }

        private static void WriteWorkbookRels(ZipArchive z, int sheetCount, bool hasStyles, bool hasSharedStrings = true)
        {
            var sb = new StringBuilder();
            sb.AppendLine(XmlHeader);
            sb.AppendLine("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (var i = 0; i < sheetCount; i++)
                sb.AppendLine($"  <Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
            var nextRid = sheetCount + 1;
            if (hasSharedStrings)
                sb.AppendLine($"  <Relationship Id=\"rId{nextRid++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            if (hasStyles)
                sb.AppendLine($"  <Relationship Id=\"rId{nextRid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.AppendLine("</Relationships>");
            AddEntry(z, "xl/_rels/workbook.xml.rels", sb.ToString());
        }

        private static void WriteWorkbook(ZipArchive z, params string[] sheetNames)
        {
            var sb = new StringBuilder();
            sb.AppendLine(XmlHeader);
            sb.AppendLine("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.AppendLine("  <sheets>");
            for (var i = 0; i < sheetNames.Length; i++)
                sb.AppendLine($"    <sheet name=\"{Esc(sheetNames[i])}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
            sb.AppendLine("  </sheets>");
            sb.AppendLine("</workbook>");
            AddEntry(z, "xl/workbook.xml", sb.ToString());
        }

        private static (string sheetXml, string ssXml) BuildSheetAndStrings(string[][] rows)
        {
            var ss = new SharedStrings();
            var sb = new StringBuilder();
            sb.AppendLine(XmlHeader);
            sb.AppendLine(WsOpen);
            sb.AppendLine("<sheetData>");
            for (var r = 0; r < rows.Length; r++)
            {
                sb.AppendLine($"  <row r=\"{r + 1}\">");
                for (var c = 0; c < rows[r].Length; c++)
                {
                    var idx = ss.Add(rows[r][c]);
                    sb.AppendLine($"    <c r=\"{ColLetter(c)}{r + 1}\" t=\"s\"><v>{idx}</v></c>");
                }
                sb.AppendLine("  </row>");
            }
            sb.AppendLine("</sheetData></worksheet>");
            return (sb.ToString(), ss.ToXml());
        }

        /// <summary>Simple shared-string accumulator.</summary>
        private sealed class SharedStrings
        {
            private readonly List<string> _list = new();
            private readonly Dictionary<string, int> _map = new();

            public int Add(string value)
            {
                if (_map.TryGetValue(value, out var idx)) return idx;
                idx = _list.Count;
                _list.Add(value);
                _map[value] = idx;
                return idx;
            }

            public string ToXml()
            {
                var sb = new StringBuilder();
                sb.AppendLine(XmlHeader);
                sb.AppendLine($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_list.Count}\" uniqueCount=\"{_list.Count}\">");
                foreach (var s in _list)
                    sb.AppendLine($"  <si><t>{Esc(s)}</t></si>");
                sb.AppendLine("</sst>");
                return sb.ToString();
            }
        }
    }
}
