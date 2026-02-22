namespace MiniPdf.Tests;

public class PdfDocumentTests
{
    [Fact]
    public void AddPage_DefaultSize_CreatesUsLetterPage()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();

        Assert.Single(doc.Pages);
        Assert.Equal(612, page.Width);
        Assert.Equal(792, page.Height);
    }

    [Fact]
    public void AddPage_CustomSize_UsesProvidedDimensions()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage(width: 100, height: 200);

        Assert.Equal(100, page.Width);
        Assert.Equal(200, page.Height);
    }

    [Fact]
    public void AddText_StoresTextBlock()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddText("Hello", 10, 20, 14);

        Assert.Single(page.TextBlocks);
        var block = page.TextBlocks[0];
        Assert.Equal("Hello", block.Text);
        Assert.Equal(10, block.X);
        Assert.Equal(20, block.Y);
        Assert.Equal(14, block.FontSize);
    }

    [Fact]
    public void AddText_Chaining_ReturnsSamePage()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        var result = page.AddText("A", 0, 0).AddText("B", 0, 0);

        Assert.Same(page, result);
        Assert.Equal(2, page.TextBlocks.Count);
    }

    [Fact]
    public void Save_ProducesValidPdfHeader()
    {
        var doc = new PdfDocument();
        doc.AddPage().AddText("Test", 50, 700);

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.StartsWith("%PDF-1.4", content);
        Assert.Contains("%%EOF", content);
    }

    [Fact]
    public void Save_ContainsTextContent()
    {
        var doc = new PdfDocument();
        doc.AddPage().AddText("Hello World", 50, 700);

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("Hello World", content);
        Assert.Contains("/F1", content);
        Assert.Contains("/Helvetica", content);
    }

    [Fact]
    public void Save_MultiplePages_AllIncluded()
    {
        var doc = new PdfDocument();
        doc.AddPage().AddText("Page 1", 50, 700);
        doc.AddPage().AddText("Page 2", 50, 700);

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("Page 1", content);
        Assert.Contains("Page 2", content);
        Assert.Contains("/Count 2", content);
    }

    [Fact]
    public void Save_ToFile_CreatesFile()
    {
        var doc = new PdfDocument();
        doc.AddPage().AddText("File test", 50, 700);

        var path = Path.Combine(Path.GetTempPath(), $"minipdf_test_{Guid.NewGuid()}.pdf");
        try
        {
            doc.Save(path);
            Assert.True(File.Exists(path));
            var bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length > 0);
            Assert.StartsWith("%PDF-1.4", System.Text.Encoding.ASCII.GetString(bytes));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Save_EscapesSpecialCharacters()
    {
        var doc = new PdfDocument();
        doc.AddPage().AddText("Hello (world) \\ test", 50, 700);

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("Hello \\(world\\) \\\\ test", content);
    }

    [Fact]
    public void AddTextWrapped_WrapsLongText()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        var longText = "This is a very long text that should be wrapped across multiple lines when rendered on the page";
        page.AddTextWrapped(longText, 50, 700, maxWidth: 200, fontSize: 12);

        // Should have created multiple text blocks
        Assert.True(page.TextBlocks.Count > 1, "Long text should wrap into multiple lines");
    }

    [Fact]
    public void AddTextWrapped_EmptyText_DoesNothing()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddTextWrapped("", 50, 700, maxWidth: 200);

        Assert.Empty(page.TextBlocks);
    }

    [Fact]
    public void EmptyDocument_ProducesValidPdf()
    {
        var doc = new PdfDocument();
        doc.AddPage(); // Empty page

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.StartsWith("%PDF-1.4", content);
        Assert.Contains("%%EOF", content);
        Assert.Contains("/Type /Page", content);
    }

    [Fact]
    public void AddPage_ZeroWidth_Throws()
    {
        var doc = new PdfDocument();
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddPage(width: 0, height: 100));
    }

    [Fact]
    public void AddPage_NegativeHeight_Throws()
    {
        var doc = new PdfDocument();
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddPage(width: 100, height: -1));
    }

    [Fact]
    public void Save_NullFilePath_Throws()
    {
        var doc = new PdfDocument();
        doc.AddPage();
        Assert.Throws<ArgumentNullException>(() => doc.Save((string)null!));
    }

    [Fact]
    public void Save_EmptyFilePath_Throws()
    {
        var doc = new PdfDocument();
        doc.AddPage();
        Assert.Throws<ArgumentException>(() => doc.Save(""));
    }

    [Fact]
    public void Save_NullStream_Throws()
    {
        var doc = new PdfDocument();
        doc.AddPage();
        Assert.Throws<ArgumentNullException>(() => doc.Save((Stream)null!));
    }

    [Fact]
    public void AddText_NullText_Throws()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        Assert.Throws<ArgumentNullException>(() => page.AddText(null!, 0, 0));
    }

    [Fact]
    public void AddTextWrapped_NullText_Throws()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        Assert.Throws<ArgumentNullException>(() => page.AddTextWrapped(null!, 0, 0, 100));
    }

    [Fact]
    public void AddTextWrapped_ZeroMaxWidth_Throws()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        Assert.Throws<ArgumentOutOfRangeException>(() => page.AddTextWrapped("text", 0, 0, 0));
    }

    [Fact]
    public void Metadata_Title_IncludedInPdf()
    {
        var doc = new PdfDocument();
        doc.Title = "Test Title";
        doc.AddPage();

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Title (Test Title)", content);
    }

    [Fact]
    public void Metadata_AllProperties_IncludedInPdf()
    {
        var doc = new PdfDocument();
        doc.Title = "My Title";
        doc.Author = "My Author";
        doc.Subject = "My Subject";
        doc.Keywords = "My Keywords";
        doc.Creator = "My Creator";
        doc.AddPage();

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Title (My Title)", content);
        Assert.Contains("/Author (My Author)", content);
        Assert.Contains("/Subject (My Subject)", content);
        Assert.Contains("/Keywords (My Keywords)", content);
        Assert.Contains("/Creator (My Creator)", content);
        Assert.Contains("/Info", content);
    }

    [Fact]
    public void Metadata_None_NoInfoDictionary()
    {
        var doc = new PdfDocument();
        doc.AddPage();

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain("/Info", content);
    }

    [Fact]
    public void Metadata_SpecialChars_Escaped()
    {
        var doc = new PdfDocument();
        doc.Title = "Hello (World)";
        doc.AddPage();

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Title (Hello \\(World\\))", content);
    }
}
