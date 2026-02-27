namespace MiniPdf;

/// <summary>
/// Represents a PDF document that can contain pages with text content.
/// </summary>
internal sealed class PdfDocument
{
    private readonly List<PdfPage> _pages = [];

    /// <summary>
    /// Gets the pages in this document.
    /// </summary>
    public IReadOnlyList<PdfPage> Pages => _pages;

    /// <summary>
    /// Adds a new page to the document.
    /// </summary>
    /// <param name="width">Page width in points (default: 612 = US Letter).</param>
    /// <param name="height">Page height in points (default: 792 = US Letter).</param>
    /// <returns>The newly created page.</returns>
    public PdfPage AddPage(float width = 612, float height = 792)
    {
        var page = new PdfPage(width, height);
        _pages.Add(page);
        return page;
    }

    /// <summary>
    /// Saves the PDF document to a file.
    /// </summary>
    public void Save(string filePath)
    {
        using var stream = File.Create(filePath);
        Save(stream);
    }

    /// <summary>
    /// Saves the PDF document to a stream.
    /// </summary>
    public void Save(Stream stream)
    {
        var writer = new PdfWriter(stream);
        writer.Write(this);
    }

    /// <summary>
    /// Saves the PDF document to a byte array.
    /// </summary>
    public byte[] ToArray()
    {
        using var ms = new MemoryStream();
        Save(ms);
        return ms.ToArray();
    }
}
