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

        foreach (var sheet in sheets)
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

        var maxCols = sheet.Rows.Max(r => r.Count);
        if (maxCols == 0)
        {
            // All rows are empty — still render an empty page worth of vertical space
            doc.AddPage(options.PageWidth, options.PageHeight);
            return;
        }

        var pageWidth = options.PageWidth;
        var pageHeight = options.PageHeight;
        var usableWidth = pageWidth - options.MarginLeft - options.MarginRight;
        var avgCharWidth = options.FontSize * 0.5f;

        // Determine column widths first to decide on layout strategy
        var columnPadding = options.ColumnPadding;
        if (maxCols > 6)
        {
            columnPadding = Math.Max(4f, options.ColumnPadding * 6f / maxCols);
        }

        var colWidths = CalculateColumnWidths(sheet, maxCols, usableWidth, options, columnPadding);

        // Check if all columns fit on one page, or need to be split into groups
        var totalTableWidth = colWidths.Sum() + columnPadding * (maxCols - 1);

        if (totalTableWidth > usableWidth * 1.1f && maxCols > 3)
        {
            // Columns don't fit — split into groups that fit on a page each
            RenderSheetColumnGroups(doc, sheet, options, pageWidth, pageHeight, usableWidth, columnPadding, avgCharWidth);
        }
        else
        {
            // Single group — render normally
            RenderSheetRows(doc, sheet, options, pageWidth, pageHeight, Enumerable.Range(0, maxCols).ToArray(), columnPadding, colWidths, avgCharWidth);
        }
    }

    /// <summary>
    /// Split columns into groups that fit within usable width, render each group on separate pages.
    /// </summary>
    private static void RenderSheetColumnGroups(PdfDocument doc, ExcelSheet sheet, ConversionOptions options,
        float pageWidth, float pageHeight, float usableWidth, float columnPadding, float avgCharWidth)
    {
        var maxCols = sheet.Rows.Max(r => r.Count);

        // Calculate individual column widths without scaling
        var colMaxLengths = new int[maxCols];
        foreach (var row in sheet.Rows)
        {
            for (var col = 0; col < row.Count && col < maxCols; col++)
            {
                colMaxLengths[col] = Math.Max(colMaxLengths[col], row[col].Text.Length);
            }
        }

        // Cap each column width to MaxColumnWidth
        var maxColWidth = usableWidth * 0.6f; // No single column takes more than 60% of page
        var naturalWidths = new float[maxCols];
        for (var i = 0; i < maxCols; i++)
        {
            naturalWidths[i] = Math.Min(Math.Max(colMaxLengths[i], 3) * avgCharWidth, maxColWidth);
        }

        // Group columns to fit within usable width
        var groups = new List<int[]>();
        var currentGroup = new List<int>();
        var currentWidth = 0f;

        for (var col = 0; col < maxCols; col++)
        {
            var colWithPadding = naturalWidths[col] + (currentGroup.Count > 0 ? columnPadding : 0);
            if (currentGroup.Count > 0 && currentWidth + colWithPadding > usableWidth)
            {
                // Start new group
                groups.Add(currentGroup.ToArray());
                currentGroup = new List<int> { col };
                currentWidth = naturalWidths[col];
            }
            else
            {
                currentGroup.Add(col);
                currentWidth += colWithPadding;
            }
        }
        if (currentGroup.Count > 0) groups.Add(currentGroup.ToArray());

        // Render each column group
        foreach (var group in groups)
        {
            // Calculate column widths for this group
            var groupWidths = new float[group.Length];
            for (var i = 0; i < group.Length; i++)
            {
                groupWidths[i] = naturalWidths[group[i]];
            }

            // Scale to fit if needed
            var groupTotalWidth = groupWidths.Sum() + columnPadding * (group.Length - 1);
            if (groupTotalWidth > usableWidth)
            {
                var available = usableWidth - columnPadding * (group.Length - 1);
                var scale = available / groupWidths.Sum();
                for (var i = 0; i < groupWidths.Length; i++)
                {
                    groupWidths[i] = Math.Max(groupWidths[i] * scale, avgCharWidth);
                }
            }

            RenderSheetRows(doc, sheet, options, pageWidth, pageHeight, group, columnPadding, groupWidths, avgCharWidth);
        }
    }

    /// <summary>
    /// Render rows for a specific set of columns.
    /// </summary>
    private static void RenderSheetRows(PdfDocument doc, ExcelSheet sheet, ConversionOptions options,
        float pageWidth, float pageHeight, int[] columns, float columnPadding, float[] colWidths, float avgCharWidth)
    {
        var lineHeight = options.FontSize * options.LineSpacing;
        PdfPage? currentPage = null;
        var currentY = 0f;

        void EnsurePage()
        {
            if (currentPage == null || currentY < options.MarginBottom)
            {
                currentPage = doc.AddPage(pageWidth, pageHeight);
                currentY = pageHeight - options.MarginTop;
            }
        }

        // Sheet header (only for first column group, skip generic names like Sheet1)
        if (columns[0] == 0 && options.IncludeSheetName && !string.IsNullOrEmpty(sheet.Name) && !IsDefaultSheetName(sheet.Name))
        {
            EnsurePage();
            currentPage!.AddText(sheet.Name, options.MarginLeft, currentY, options.FontSize + 4);
            currentY -= lineHeight * 1.5f;
        }

        // Render rows
        foreach (var row in sheet.Rows)
        {
            EnsurePage();

            if (row.Count == 0)
            {
                // Empty row (sparse gap) — still advance Y
                currentY -= lineHeight;
                continue;
            }

            // Calculate wrapped lines for each column in this group
            var maxLinesInRow = 1;
            var cellLines = new string[columns.Length][];

            for (var i = 0; i < columns.Length; i++)
            {
                var col = columns[i];
                if (col < row.Count)
                {
                    var cellText = row[col].Text;
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        var maxChars = Math.Max(1, (int)(colWidths[i] / avgCharWidth));
                        var wrapped = WrapCellText(cellText, maxChars);
                        cellLines[i] = wrapped;
                        if (wrapped.Length > maxLinesInRow) maxLinesInRow = wrapped.Length;
                    }
                    else
                    {
                        cellLines[i] = Array.Empty<string>();
                    }
                }
                else
                {
                    cellLines[i] = Array.Empty<string>();
                }
            }

            // Check space for wrapped lines
            var rowHeight = lineHeight * maxLinesInRow;
            if (currentY - rowHeight < options.MarginBottom && currentPage != null)
            {
                currentPage = doc.AddPage(pageWidth, pageHeight);
                currentY = pageHeight - options.MarginTop;
            }

            // Render cells
            var x = options.MarginLeft;
            for (var i = 0; i < columns.Length; i++)
            {
                var lines = cellLines[i];
                var col = columns[i];
                var color = col < row.Count ? row[col].Color : null;
                var cellY = currentY;

                for (var lineIdx = 0; lineIdx < lines.Length; lineIdx++)
                {
                    if (!string.IsNullOrEmpty(lines[lineIdx]))
                    {
                        currentPage!.AddText(lines[lineIdx], x, cellY, options.FontSize, color);
                    }
                    cellY -= lineHeight;
                }

                x += colWidths[i] + columnPadding;
            }

            currentY -= rowHeight;
        }
    }

    /// <summary>
    /// Wrap a single cell text into multiple lines at word boundaries.
    /// </summary>
    private static string[] WrapCellText(string text, int maxCharsPerLine)
    {
        if (maxCharsPerLine <= 0) maxCharsPerLine = 1;
        if (text.Length <= maxCharsPerLine) return new[] { text };

        var lines = new List<string>();
        var words = text.Split(' ');

        var currentLine = "";
        foreach (var word in words)
        {
            if (currentLine.Length == 0)
            {
                currentLine = word;
            }
            else if (currentLine.Length + 1 + word.Length <= maxCharsPerLine)
            {
                currentLine += " " + word;
            }
            else
            {
                // If current line overflows, hard-break it
                while (currentLine.Length > maxCharsPerLine)
                {
                    lines.Add(currentLine[..maxCharsPerLine]);
                    currentLine = currentLine[maxCharsPerLine..];
                }
                if (currentLine.Length > 0)
                    lines.Add(currentLine);
                currentLine = word;
            }
        }

        // Handle the last line — might also need hard-breaking
        while (currentLine.Length > maxCharsPerLine)
        {
            lines.Add(currentLine[..maxCharsPerLine]);
            currentLine = currentLine[maxCharsPerLine..];
        }
        if (currentLine.Length > 0)
            lines.Add(currentLine);

        return lines.ToArray();
    }

    /// <summary>
    /// Checks if a sheet name is a generic default like Sheet1, Sheet2, etc.
    /// </summary>
    private static bool IsDefaultSheetName(string name)
    {
        if (name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase) && name.Length <= 8)
        {
            return int.TryParse(name.AsSpan(5), out _);
        }
        return false;
    }

    private static float[] CalculateColumnWidths(ExcelSheet sheet, int maxCols, float usableWidth, ConversionOptions options, float columnPadding)
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

        // Cap max column width to prevent very long text from consuming all width.
        // Only apply aggressive cap when text is truly long (>80 chars).
        // For normal text, allow up to 60% of usable width.
        var maxColWidth = usableWidth * 0.6f;

        // Calculate natural widths with cap
        var naturalWidths = new float[maxCols];
        for (var i = 0; i < maxCols; i++)
        {
            var naturalWidth = Math.Max(colMaxLengths[i], 3) * avgCharWidth;
            naturalWidths[i] = Math.Min(naturalWidth, maxColWidth);
        }
        var totalPadding = columnPadding * (maxCols - 1);
        var totalNatural = naturalWidths.Sum() + totalPadding;

        // Scale down if exceeding usable width
        if (totalNatural > usableWidth)
        {
            var availableForColumns = usableWidth - totalPadding;
            if (availableForColumns <= 0)
            {
                // Extreme case: even padding alone exceeds width; distribute evenly
                availableForColumns = usableWidth * 0.9f;
            }
            var scale = availableForColumns / naturalWidths.Sum();
            for (var i = 0; i < naturalWidths.Length; i++)
            {
                naturalWidths[i] = Math.Max(naturalWidths[i] * scale, avgCharWidth);
            }
        }

        return naturalWidths;
    }
}
