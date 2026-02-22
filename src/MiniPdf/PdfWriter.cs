using System.Globalization;
using System.Text;

namespace MiniPdf;

/// <summary>
/// Low-level PDF writer. Produces valid PDF 1.4 output with Helvetica font.
/// </summary>
internal sealed class PdfWriter
{
    private readonly Stream _stream;
    private readonly List<long> _objectOffsets = [];
    private int _objectCount;

    internal PdfWriter(Stream stream)
    {
        _stream = stream;
    }

    internal void Write(PdfDocument document)
    {
        // PDF Header
        WriteRaw("%PDF-1.4\n");
        // Binary comment to signal binary content (recommended by spec)
        WriteRaw("%\xe2\xe3\xcf\xd3\n");

        // Build object tree:
        // Obj 1: Catalog
        // Obj 2: Pages
        // Obj 3: Font (Helvetica)
        // Obj 4+: Page objects and their content streams

        var pageObjectNumbers = new List<int>();
        var pageContentPairs = new List<(int pageObj, int contentObj)>();

        // Reserve objects 1 (Catalog), 2 (Pages), 3 (Font)
        // Then allocate page + content stream pairs
        var nextObj = 4;
        foreach (var _ in document.Pages)
        {
            var pageObj = nextObj++;
            var contentObj = nextObj++;
            pageObjectNumbers.Add(pageObj);
            pageContentPairs.Add((pageObj, contentObj));
        }

        _objectCount = nextObj - 1;
        _objectOffsets.Clear();
        // Pad index 0 (not used since PDF objects are 1-based)
        for (var i = 0; i < nextObj; i++)
            _objectOffsets.Add(0);

        // Write Catalog (object 1)
        _objectOffsets[1] = Position;
        WriteRaw("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

        // Write Pages (object 2)
        _objectOffsets[2] = Position;
        var kids = string.Join(" ", pageObjectNumbers.Select(n => $"{n} 0 R"));
        WriteRaw($"2 0 obj\n<< /Type /Pages /Kids [{kids}] /Count {document.Pages.Count} >>\nendobj\n");

        // Write Font (object 3) - Helvetica (built-in, no embedding needed)
        _objectOffsets[3] = Position;
        WriteRaw("3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj\n");

        // Write each page and its content stream
        for (var i = 0; i < document.Pages.Count; i++)
        {
            var page = document.Pages[i];
            var (pageObj, contentObj) = pageContentPairs[i];

            // Build content stream
            var content = BuildContentStream(page);
            var contentBytes = Encoding.ASCII.GetBytes(content);

            // Write content stream object
            _objectOffsets[contentObj] = Position;
            WriteRaw($"{contentObj} 0 obj\n<< /Length {contentBytes.Length} >>\nstream\n");
            _stream.Write(contentBytes);
            WriteRaw("\nendstream\nendobj\n");

            // Write page object
            var w = page.Width.ToString(CultureInfo.InvariantCulture);
            var h = page.Height.ToString(CultureInfo.InvariantCulture);
            _objectOffsets[pageObj] = Position;
            WriteRaw($"{pageObj} 0 obj\n");
            WriteRaw($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {w} {h}] ");
            WriteRaw($"/Contents {contentObj} 0 R /Resources << /Font << /F1 3 0 R >> >> >>\n");
            WriteRaw("endobj\n");
        }

        // Write xref table
        var xrefOffset = Position;
        WriteRaw("xref\n");
        WriteRaw($"0 {_objectCount + 1}\n");
        WriteRaw("0000000000 65535 f \n");
        for (var i = 1; i <= _objectCount; i++)
        {
            WriteRaw($"{_objectOffsets[i]:D10} 00000 n \n");
        }

        // Write trailer
        WriteRaw("trailer\n");
        WriteRaw($"<< /Size {_objectCount + 1} /Root 1 0 R >>\n");
        WriteRaw("startxref\n");
        WriteRaw($"{xrefOffset}\n");
        WriteRaw("%%EOF\n");

        _stream.Flush();
    }

    private static string BuildContentStream(PdfPage page)
    {
        var sb = new StringBuilder();
        sb.Append("BT\n");

        foreach (var block in page.TextBlocks)
        {
            var fontSize = block.FontSize.ToString(CultureInfo.InvariantCulture);
            var x = block.X.ToString(CultureInfo.InvariantCulture);
            var y = block.Y.ToString(CultureInfo.InvariantCulture);
            var escapedText = EscapePdfString(block.Text);

            // Set text color if not black
            if (!block.Color.IsBlack)
            {
                var r = block.Color.R.ToString("F3", CultureInfo.InvariantCulture);
                var g = block.Color.G.ToString("F3", CultureInfo.InvariantCulture);
                var b = block.Color.B.ToString("F3", CultureInfo.InvariantCulture);
                sb.Append($"{r} {g} {b} rg\n");
            }
            else
            {
                sb.Append("0 0 0 rg\n");
            }

            sb.Append($"/F1 {fontSize} Tf\n");
            sb.Append($"{x} {y} Td\n");
            sb.Append($"({escapedText}) Tj\n");
            // Reset position for next absolute placement
            var nx = (-block.X).ToString(CultureInfo.InvariantCulture);
            var ny = (-block.Y).ToString(CultureInfo.InvariantCulture);
            sb.Append($"{nx} {ny} Td\n");
        }

        sb.Append("ET\n");
        return sb.ToString();
    }

    private static string EscapePdfString(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("(", "\\(")
            .Replace(")", "\\)")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n");
    }

    private long Position => _stream.Position;

    private void WriteRaw(string text)
    {
        var bytes = Encoding.ASCII.GetBytes(text);
        _stream.Write(bytes);
    }
}
