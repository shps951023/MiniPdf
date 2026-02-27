using System.Globalization;

namespace MiniPdf;

/// <summary>
/// Converts Excel (.xlsx) files to PDF documents.
/// Renders cell text in a simple table layout using the built-in Helvetica font.
/// </summary>
internal static class ExcelToPdfConverter
{
    /// <summary>
    /// Options for controlling Excel-to-PDF conversion.
    /// </summary>
    internal sealed class ConversionOptions
    {
        /// <summary>Font size in points (default: 10).</summary>
        public float FontSize { get; set; } = 10;

        /// <summary>Page left margin in points (default: 50).</summary>
        public float MarginLeft { get; set; } = 50;

        /// <summary>Page top margin in points (default: 50).</summary>
        public float MarginTop { get; set; } = 50;

        /// <summary>Page right margin in points (default: 50).</summary>
        public float MarginRight { get; set; } = 50;

        /// <summary>Page bottom margin in points (default: 50).</summary>
        public float MarginBottom { get; set; } = 50;

        /// <summary>Padding between columns in points (default: 20).</summary>
        public float ColumnPadding { get; set; } = 20;

        /// <summary>Line spacing multiplier (default: 1.4).</summary>
        public float LineSpacing { get; set; } = 1.4f;

        /// <summary>Page width in points (default: 612 = US Letter).</summary>
        public float PageWidth { get; set; } = 612;

        /// <summary>Page height in points (default: 792 = US Letter).</summary>
        public float PageHeight { get; set; } = 792;

        /// <summary>Whether to include sheet name as a header (default: true).</summary>
        public bool IncludeSheetName { get; set; } = true;

        /// <summary>
        /// Zero-based indices of the sheets to include. When null or empty, all sheets are included.
        /// </summary>
        public IReadOnlyList<int>? SheetIndices { get; set; }

        /// <summary>
        /// Names of the sheets to include (case-insensitive). When null or empty, all sheets are included.
        /// Ignored if <see cref="SheetIndices"/> is also specified.
        /// </summary>
        public IReadOnlyList<string>? SheetNames { get; set; }
    }

    /// <summary>
    /// Converts an Excel file to a PDF document.
    /// </summary>
    /// <param name="excelPath">Path to the .xlsx file.</param>
    /// <param name="options">Optional conversion settings.</param>
    /// <returns>A PdfDocument containing the Excel data.</returns>
    internal static PdfDocument Convert(string excelPath, ConversionOptions? options = null)
    {
        using var stream = File.OpenRead(excelPath);
        return Convert(stream, options);
    }

    /// <summary>
    /// Converts an Excel stream to a PDF document.
    /// </summary>
    /// <param name="excelStream">Stream containing .xlsx data.</param>
    /// <param name="options">Optional conversion settings.</param>
    /// <returns>A PdfDocument containing the Excel data.</returns>
    internal static PdfDocument Convert(Stream excelStream, ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        var sheets = ExcelReader.ReadSheets(excelStream);
        var doc = new PdfDocument();

        // Filter sheets by index or name if specified
        IEnumerable<ExcelSheet> sheetsToRender = sheets;
        if (options.SheetIndices is { Count: > 0 })
        {
            var indexSet = new HashSet<int>(options.SheetIndices);
            sheetsToRender = sheets.Where((_, i) => indexSet.Contains(i));
        }
        else if (options.SheetNames is { Count: > 0 })
        {
            var nameSet = new HashSet<string>(options.SheetNames, StringComparer.OrdinalIgnoreCase);
            sheetsToRender = sheets.Where(s => nameSet.Contains(s.Name));
        }

        foreach (var sheet in sheetsToRender)
        {
            RenderSheet(doc, sheet, options);
        }

        // If no sheets found, create at least one empty page
        if (doc.Pages.Count == 0)
        {
            doc.AddPage(options.PageWidth, options.PageHeight);
        }

        return doc;
    }

    /// <summary>
    /// Converts an Excel file directly to a PDF file.
    /// </summary>
    /// <param name="excelPath">Path to the .xlsx file.</param>
    /// <param name="pdfPath">Path for the output .pdf file.</param>
    /// <param name="options">Optional conversion settings.</param>
    internal static void ConvertToFile(string excelPath, string pdfPath, ConversionOptions? options = null)
    {
        var doc = Convert(excelPath, options);
        doc.Save(pdfPath);
    }

    private static void RenderSheet(PdfDocument doc, ExcelSheet sheet, ConversionOptions options)
    {
        if (sheet.Rows.Count == 0) return;

        var lineHeight = options.FontSize * options.LineSpacing;
        var usableWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        var usableHeight = options.PageHeight - options.MarginTop - options.MarginBottom;

        // Calculate column widths based on content
        var maxCols = sheet.Rows.Max(r => r.Count);
        if (maxCols == 0) return;

        var colWidths = CalculateColumnWidths(sheet, maxCols, usableWidth, options);

        // Render rows across pages
        PdfPage? currentPage = null;
        var currentY = 0f;

        void EnsurePage()
        {
            if (currentPage == null || currentY < options.MarginBottom)
            {
                currentPage = doc.AddPage(options.PageWidth, options.PageHeight);
                currentY = options.PageHeight - options.MarginTop;
            }
        }

        // Sheet header
        if (options.IncludeSheetName && !string.IsNullOrEmpty(sheet.Name))
        {
            EnsurePage();
            currentPage!.AddText(sheet.Name, options.MarginLeft, currentY, options.FontSize + 4);
            currentY -= lineHeight * 1.5f;
        }

        // Render rows
        foreach (var row in sheet.Rows)
        {
            EnsurePage();

            var x = options.MarginLeft;
            for (var col = 0; col < maxCols && col < row.Count; col++)
            {
                var cell = row[col];
                var cellText = cell.Text;
                if (!string.IsNullOrEmpty(cellText))
                {
                    // Truncate text if it exceeds column width
                    var maxChars = (int)(colWidths[col] / (options.FontSize * 0.5f));
                    if (cellText.Length > maxChars && maxChars > 3)
                    {
                        cellText = cellText[..(maxChars - 3)] + "...";
                    }

                    currentPage!.AddText(cellText, x, currentY, options.FontSize, cell.Color);
                }
                x += colWidths[col] + options.ColumnPadding;
            }

            currentY -= lineHeight;
        }
    }

    private static float[] CalculateColumnWidths(ExcelSheet sheet, int maxCols, float usableWidth, ConversionOptions options)
    {
        var avgCharWidth = options.FontSize * 0.5f;
        var colMaxLengths = new int[maxCols];

        foreach (var row in sheet.Rows)
        {
            for (var col = 0; col < row.Count && col < maxCols; col++)
            {
                colMaxLengths[col] = Math.Max(colMaxLengths[col], row[col].Text.Length);
            }
        }

        // Calculate natural widths
        var naturalWidths = colMaxLengths.Select(len => Math.Max(len, 3) * avgCharWidth).ToArray();
        var totalNatural = naturalWidths.Sum() + options.ColumnPadding * (maxCols - 1);

        // Scale down if exceeding usable width
        if (totalNatural > usableWidth)
        {
            var availableForColumns = usableWidth - options.ColumnPadding * (maxCols - 1);
            var scale = availableForColumns / naturalWidths.Sum();
            for (var i = 0; i < naturalWidths.Length; i++)
            {
                naturalWidths[i] *= scale;
            }
        }

        return naturalWidths;
    }
}
