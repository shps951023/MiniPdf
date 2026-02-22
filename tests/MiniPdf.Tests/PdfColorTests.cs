namespace MiniPdf.Tests;

public class PdfColorTests
{
    [Fact]
    public void FromRgb_CreatesCorrectColor()
    {
        var color = PdfColor.FromRgb(255, 0, 0);
        Assert.Equal(1f, color.R);
        Assert.Equal(0f, color.G);
        Assert.Equal(0f, color.B);
    }

    [Fact]
    public void FromHex_6Char_ParsesCorrectly()
    {
        var color = PdfColor.FromHex("00FF00");
        Assert.Equal(0f, color.R);
        Assert.Equal(1f, color.G);
        Assert.Equal(0f, color.B);
    }

    [Fact]
    public void FromHex_WithHash_ParsesCorrectly()
    {
        var color = PdfColor.FromHex("#0000FF");
        Assert.Equal(0f, color.R);
        Assert.Equal(0f, color.G);
        Assert.Equal(1f, color.B);
    }

    [Fact]
    public void FromHex_8CharArgb_SkipsAlpha()
    {
        var color = PdfColor.FromHex("FFFF0000"); // Alpha=FF, R=FF, G=00, B=00
        Assert.Equal(1f, color.R);
        Assert.Equal(0f, color.G);
        Assert.Equal(0f, color.B);
    }

    [Fact]
    public void FromHex_Invalid_ReturnsBlack()
    {
        var color = PdfColor.FromHex("xyz");
        Assert.Equal(PdfColor.Black, color);
    }

    [Fact]
    public void FromHex_Empty_ReturnsBlack()
    {
        var color = PdfColor.FromHex("");
        Assert.Equal(PdfColor.Black, color);
    }

    [Fact]
    public void IsBlack_TrueForBlack()
    {
        Assert.True(PdfColor.Black.IsBlack);
        Assert.True(new PdfColor(0, 0, 0).IsBlack);
    }

    [Fact]
    public void IsBlack_FalseForNonBlack()
    {
        Assert.False(PdfColor.Red.IsBlack);
        Assert.False(PdfColor.Blue.IsBlack);
    }

    [Fact]
    public void Equality_SameValues_AreEqual()
    {
        var a = PdfColor.FromRgb(128, 64, 32);
        var b = PdfColor.FromRgb(128, 64, 32);
        Assert.Equal(a, b);
        Assert.True(a == b);
    }

    [Fact]
    public void Equality_DifferentValues_AreNotEqual()
    {
        Assert.NotEqual(PdfColor.Red, PdfColor.Blue);
        Assert.True(PdfColor.Red != PdfColor.Blue);
    }

    [Fact]
    public void Constructor_ClampsValues()
    {
        var color = new PdfColor(1.5f, -0.5f, 0.5f);
        Assert.Equal(1f, color.R);
        Assert.Equal(0f, color.G);
        Assert.Equal(0.5f, color.B);
    }

    [Fact]
    public void AddText_WithColor_StoresColor()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddText("Red text", 50, 700, 12, PdfColor.Red);

        Assert.Single(page.TextBlocks);
        Assert.Equal(PdfColor.Red, page.TextBlocks[0].Color);
    }

    [Fact]
    public void AddText_WithoutColor_DefaultsToBlack()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddText("Default text", 50, 700);

        Assert.Equal(PdfColor.Black, page.TextBlocks[0].Color);
    }

    [Fact]
    public void Save_WithColor_ContainsRgOperator()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddText("Red text", 50, 700, 12, PdfColor.Red);

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        // Red = 1.000 0.000 0.000 rg
        Assert.Contains("1.000 0.000 0.000 rg", content);
        Assert.Contains("Red text", content);
    }

    [Fact]
    public void Save_MixedColors_AllPresent()
    {
        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.AddText("Red", 50, 700, 12, PdfColor.Red);
        page.AddText("Blue", 50, 680, 12, PdfColor.Blue);
        page.AddText("Black", 50, 660, 12); // default

        var bytes = doc.ToArray();
        var content = System.Text.Encoding.ASCII.GetString(bytes);

        Assert.Contains("1.000 0.000 0.000 rg", content);
        Assert.Contains("0.000 0.000 1.000 rg", content);
        Assert.Contains("0 0 0 rg", content);
    }
}
