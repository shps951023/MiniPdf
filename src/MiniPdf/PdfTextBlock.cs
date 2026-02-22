namespace MiniPdf;

/// <summary>
/// Represents a text block to be rendered on a PDF page.
/// </summary>
public sealed class PdfTextBlock
{
    /// <summary>
    /// The text content.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// X position in points from the left edge.
    /// </summary>
    public float X { get; }

    /// <summary>
    /// Y position in points from the bottom edge.
    /// </summary>
    public float Y { get; }

    /// <summary>
    /// Font size in points.
    /// </summary>
    public float FontSize { get; }

    internal PdfTextBlock(string text, float x, float y, float fontSize)
    {
        Text = text;
        X = x;
        Y = y;
        FontSize = fontSize;
    }
}
