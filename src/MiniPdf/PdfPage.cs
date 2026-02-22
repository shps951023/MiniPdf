namespace MiniPdf;

/// <summary>
/// Represents a single page in a PDF document.
/// </summary>
public sealed class PdfPage
{
    private readonly List<PdfTextBlock> _textBlocks = [];

    /// <summary>
    /// Page width in points.
    /// </summary>
    public float Width { get; }

    /// <summary>
    /// Page height in points.
    /// </summary>
    public float Height { get; }

    /// <summary>
    /// Gets the text blocks on this page.
    /// </summary>
    public IReadOnlyList<PdfTextBlock> TextBlocks => _textBlocks;

    internal PdfPage(float width, float height)
    {
        Width = width;
        Height = height;
    }

    /// <summary>
    /// Adds a text block at the specified position.
    /// </summary>
    /// <param name="text">The text to render.</param>
    /// <param name="x">X position in points from the left edge.</param>
    /// <param name="y">Y position in points from the bottom edge.</param>
    /// <param name="fontSize">Font size in points (default: 12).</param>
    /// <returns>The current page for chaining.</returns>
    public PdfPage AddText(string text, float x, float y, float fontSize = 12)
    {
        _textBlocks.Add(new PdfTextBlock(text, x, y, fontSize));
        return this;
    }

    /// <summary>
    /// Adds text that automatically wraps within the specified region.
    /// Text flows from top to bottom, left to right within the given bounds.
    /// </summary>
    /// <param name="text">The text to render.</param>
    /// <param name="x">X position of the left edge.</param>
    /// <param name="y">Y position of the top edge.</param>
    /// <param name="maxWidth">Maximum width for text wrapping.</param>
    /// <param name="fontSize">Font size in points (default: 12).</param>
    /// <param name="lineSpacing">Line spacing multiplier (default: 1.2).</param>
    /// <returns>The current page for chaining.</returns>
    public PdfPage AddTextWrapped(string text, float x, float y, float maxWidth, float fontSize = 12, float lineSpacing = 1.2f)
    {
        if (string.IsNullOrEmpty(text))
            return this;

        var lineHeight = fontSize * lineSpacing;
        // Approximate character width for Helvetica at given font size
        var avgCharWidth = fontSize * 0.5f;
        var charsPerLine = (int)(maxWidth / avgCharWidth);
        if (charsPerLine < 1) charsPerLine = 1;

        var lines = WrapText(text, charsPerLine);
        var currentY = y;

        foreach (var line in lines)
        {
            // PDF y-coordinate is from bottom, so subtract to go down
            AddText(line, x, currentY, fontSize);
            currentY -= lineHeight;
        }

        return this;
    }

    private static List<string> WrapText(string text, int maxCharsPerLine)
    {
        var result = new List<string>();
        var paragraphs = text.Split('\n');

        foreach (var paragraph in paragraphs)
        {
            if (string.IsNullOrEmpty(paragraph))
            {
                result.Add(string.Empty);
                continue;
            }

            var words = paragraph.Split(' ', StringSplitOptions.RemoveEmptyEntries);
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
                    result.Add(currentLine);
                    currentLine = word;
                }
            }

            if (currentLine.Length > 0)
                result.Add(currentLine);
        }

        return result;
    }
}
