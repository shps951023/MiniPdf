# MiniPdf

A minimal, zero-dependency .NET library for generating PDF documents from text and Excel (.xlsx) files.

## Features

- **Text-to-PDF** — Create PDF documents with positioned or auto-wrapped text
- **Excel-to-PDF** — Convert `.xlsx` files to paginated PDF with automatic column layout
- **Text Color** — Per-cell font color support in both text and Excel-to-PDF conversion
- **PDF Metadata** — Set document Title, Author, Subject, Keywords, and Creator
- **Zero dependencies** — Uses only built-in .NET APIs (no external packages)
- **Valid PDF 1.4** output with Helvetica font
- **Input validation** — All public APIs validate arguments with descriptive exceptions

## Supported Feature Matrix

| Feature | Status | Notes |
|---|---|---|
| Text positioning | ✅ | Absolute (x,y) placement |
| Text wrapping | ✅ | Auto-wrap within specified width |
| Text color | ✅ | RGB color via `PdfColor` class |
| Excel cell text | ✅ | Shared strings, inline strings, numbers |
| Excel text color | ✅ | ARGB hex, indexed colors from styles.xml |
| Multi-sheet support | ✅ | All sheets rendered sequentially |
| Pagination | ✅ | Automatic page breaks for long content |
| Column auto-sizing | ✅ | Based on content width |
| Sheet name headers | ✅ | Optional, configurable |
| PDF metadata | ✅ | Title, Author, Subject, Keywords, Creator |
| Page size options | ✅ | US Letter (default), A4, or custom |
| Bold / italic | ❌ | Planned |
| Cell background color | ❌ | Planned |
| Text alignment | ❌ | Planned |
| Merged cells | ❌ | Planned |
| Images | ❌ | Planned |
| Number/date formatting | ❌ | Planned |
| Column width from Excel | ❌ | Planned |
| Hyperlinks | ❌ | Planned |

## Getting Started

### Requirements

- .NET 9.0 or later

### Install from NuGet

```bash
dotnet add package MiniPdf
```

### Build

```bash
dotnet build
```

### Run Tests

```bash
dotnet test
```

## Usage

### Simple Text PDF

```csharp
using MiniPdf;

var doc = new PdfDocument();
var page = doc.AddPage(); // US Letter size by default

page.AddText("Hello, World!", x: 50, y: 700, fontSize: 24);
page.AddText("This is MiniPdf.", x: 50, y: 670, fontSize: 12);

doc.Save("output.pdf");
```

### Auto-Wrapped Text

```csharp
var doc = new PdfDocument();
var page = doc.AddPage();

var longText = "This is a long paragraph that will automatically wrap "
             + "within the specified width boundary on the page.";

page.AddTextWrapped(longText, x: 50, y: 700, maxWidth: 500, fontSize: 12);

doc.Save("wrapped.pdf");
```

### Text with Color

```csharp
using MiniPdf;

var doc = new PdfDocument();
var page = doc.AddPage();

page.AddText("Red text", 50, 700, 12, PdfColor.Red);
page.AddText("Blue text", 50, 680, 12, PdfColor.Blue);
page.AddText("Custom color", 50, 660, 12, PdfColor.FromRgb(128, 64, 0));
page.AddText("Hex color", 50, 640, 12, PdfColor.FromHex("#FF8C00"));

doc.Save("colored.pdf");
```

### PDF Metadata

```csharp
using MiniPdf;

var doc = new PdfDocument();
doc.Title = "My Document";
doc.Author = "John Doe";
doc.Subject = "Sample PDF";
doc.Keywords = "pdf, sample, minipdf";
doc.Creator = "MiniPdf";

doc.AddPage().AddText("Hello with metadata!", 50, 700);

doc.Save("metadata.pdf");
```

### Excel to PDF

```csharp
using MiniPdf;

// One-liner: file to file
ExcelToPdfConverter.ConvertToFile("data.xlsx", "data.pdf");

// With options
var options = new ExcelToPdfConverter.ConversionOptions
{
    FontSize = 10,
    PageWidth = 595,   // A4
    PageHeight = 842,  // A4
    IncludeSheetName = true,
};

var doc = ExcelToPdfConverter.Convert("data.xlsx", options);
doc.Save("data.pdf");
```

### Save to Stream or Byte Array

```csharp
var doc = new PdfDocument();
doc.AddPage().AddText("Hello", 50, 700);

// To stream
using var stream = new MemoryStream();
doc.Save(stream);

// To byte array
byte[] bytes = doc.ToArray();
```

## Project Structure

```
MiniPdf.sln
├── src/MiniPdf/              # Library
│   ├── PdfDocument.cs        # Document model (pages + metadata)
│   ├── PdfPage.cs            # Page with text placement
│   ├── PdfTextBlock.cs       # Text block data
│   ├── PdfColor.cs           # RGB color for text rendering
│   ├── PdfWriter.cs          # PDF 1.4 binary writer
│   ├── ExcelReader.cs        # .xlsx parser (ZIP + XML)
│   └── ExcelToPdfConverter.cs# Excel-to-PDF public API
└── tests/MiniPdf.Tests/      # xUnit tests
```

## License

MIT
