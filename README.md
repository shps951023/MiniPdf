# MiniPdf

A minimal, zero-dependency .NET library for converting Excel (.xlsx) files to PDF.

> **Security**: All PRs are automatically reviewed by Copilot AI and Azure AI security scan for vulnerabilities.

## Features

- **Excel-to-PDF** — Convert `.xlsx` to paginated PDF in one line
- **Zero dependencies** — Uses only built-in .NET APIs (no external packages)
- **Valid PDF 1.4** output with Helvetica font

## Getting Started

```bash
dotnet add package MiniPdf
```

Requires **.NET 9.0** or later.

## Usage

```csharp
using MiniPdf;

// Convert Excel to PDF — that's it!
MiniPdf.ConvertToPdf("data.xlsx", "data.pdf");
```

Or get a byte array:

```csharp
byte[] pdf = MiniPdf.ConvertToPdf("data.xlsx");
```

Or from a stream:

```csharp
using var stream = File.OpenRead("data.xlsx");
byte[] pdf = MiniPdf.ConvertToPdf(stream);
```

## Build & Test

```bash
dotnet build
dotnet test
```

## License

MIT
