# MiniPdf

A minimal, zero-dependency .NET library for generating PDF documents from text and Excel (.xlsx) files.

## Project Scope

MiniPdf's purpose is to convert Excel (`.xlsx`) files to PDF and to create PDF documents programmatically from text, using only built-in .NET APIs.

### Supported Features

| Feature | Status |
|---|---|
| Text-to-PDF (positioned and auto-wrapped) | ✅ Supported |
| Excel-to-PDF (`.xlsx` to PDF) | ✅ Supported |
| Multiple worksheets (each sheet rendered in the PDF) | ✅ Supported |
| Basic cell formatting (font size, font color) | ✅ Supported |
| Page setup options (margins, page size, column padding, line spacing) | ✅ Supported |
| Input as file path or `Stream` | ✅ Supported |
| Output as file path, `Stream`, or `byte[]` | ✅ Supported |
| Zero external dependencies (built-in .NET APIs only) | ✅ Supported |
| Images in Excel cells | ❌ Not supported |
| Charts in Excel cells | ❌ Not supported |
| Cell borders and background fills | ❌ Not supported |
| Bold / italic / underline font styles | ❌ Not supported |
| Cell merges | ❌ Not supported |
| Formulas (rendered as their cached value) | ❌ Not supported |

### Target .NET Versions

| Framework | Supported |
|---|---|
| .NET 6.0 | ✅ |
| .NET 7.0 | ✅ |
| .NET 8.0 | ✅ |
| .NET 9.0 | ✅ |

> The library targets `net6.0` and `net9.0`. Any .NET version in that range (7, 8) is compatible via the `net6.0` target.

## Features

- **Text-to-PDF** — Create PDF documents with positioned or auto-wrapped text
- **Excel-to-PDF** — Convert `.xlsx` files to paginated PDF with automatic column layout
- **Zero dependencies** — Uses only built-in .NET APIs (no external packages)
- **Valid PDF 1.4** output with Helvetica font

## Getting Started

### Requirements

- .NET 6.0 or later

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
│   ├── PdfDocument.cs        # Document model
│   ├── PdfPage.cs            # Page with text placement
│   ├── PdfTextBlock.cs       # Text block data
│   ├── PdfWriter.cs          # PDF 1.4 binary writer
│   ├── ExcelReader.cs        # .xlsx parser (ZIP + XML)
│   └── ExcelToPdfConverter.cs# Excel-to-PDF public API
└── tests/MiniPdf.Tests/      # xUnit tests
```

## License

MIT
