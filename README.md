# MiniPdf

A minimal, zero-dependency .NET library for converting Excel (.xlsx) files to PDF.

> **Security**: All PRs are automatically reviewed by Copilot AI and Azure AI security scan for vulnerabilities.

## Features

- **Excel-to-PDF** — Convert `.xlsx` files to paginated PDF with automatic column layout
- **Zero dependencies** — Uses only built-in .NET APIs (no external packages)
- **Valid PDF 1.4** output

## Getting Started

### Install via NuGet

```bash
dotnet add package MiniPdf
```

### Requirements

- .NET 9.0 or later

## Usage

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

## License

MIT
