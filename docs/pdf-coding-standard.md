# PDF Coding Standard

This document defines the coding standards and conventions used in the MiniPdf project. All contributors should follow these guidelines when adding or modifying code.

## Table of Contents

- [Project Architecture](#project-architecture)
- [Naming Conventions](#naming-conventions)
- [Class Design](#class-design)
- [API Design Patterns](#api-design-patterns)
- [PDF 1.4 Specification Conventions](#pdf-14-specification-conventions)
- [Color Handling](#color-handling)
- [Excel Integration](#excel-integration)
- [Error Handling and Validation](#error-handling-and-validation)
- [Documentation](#documentation)
- [Testing](#testing)
- [Security](#security)
- [Dependencies](#dependencies)

---

## Project Architecture

```
MiniPdf.sln
├── src/MiniPdf/                  # Library source
│   ├── PdfDocument.cs            # Document model (public entry point)
│   ├── PdfPage.cs                # Page with text placement
│   ├── PdfTextBlock.cs           # Immutable text block data
│   ├── PdfColor.cs               # RGB color value type
│   ├── PdfWriter.cs              # PDF 1.4 binary writer (internal)
│   ├── ExcelReader.cs            # .xlsx parser (internal)
│   └── ExcelToPdfConverter.cs    # Excel-to-PDF public API
└── tests/MiniPdf.Tests/          # xUnit tests
    ├── PdfDocumentTests.cs       # Core document tests
    ├── PdfColorTests.cs          # Color parsing and rendering tests
    └── ExcelToPdfConverterTests.cs # Excel conversion tests
```

### Layering Rules

- **Public API layer**: `PdfDocument`, `PdfPage`, `PdfTextBlock`, `PdfColor`, `ExcelToPdfConverter` — these are the types consumers interact with.
- **Internal implementation layer**: `PdfWriter`, `ExcelReader`, `ExcelSheet`, `ExcelCell` — these are marked `internal` and must not be exposed to consumers.
- New types should follow this same separation: public for API surface, internal for implementation details.

---

## Naming Conventions

| Element          | Convention    | Example                          |
|------------------|---------------|----------------------------------|
| Namespace        | PascalCase    | `MiniPdf`                        |
| Class            | PascalCase    | `PdfDocument`, `PdfWriter`       |
| Method           | PascalCase    | `AddPage()`, `Save()`            |
| Property         | PascalCase    | `Width`, `Height`, `TextBlocks`  |
| Parameter        | camelCase     | `fontSize`, `maxWidth`           |
| Private field    | `_camelCase`  | `_pages`, `_objectOffsets`       |
| Local variable   | camelCase     | `lineHeight`, `avgCharWidth`     |
| Constant/static  | PascalCase    | `PdfColor.Black`, `PdfColor.Red` |

### Prefix Conventions

- All PDF-related public types are prefixed with `Pdf` (e.g., `PdfDocument`, `PdfPage`, `PdfColor`, `PdfTextBlock`).
- All Excel-related internal types are prefixed with `Excel` (e.g., `ExcelReader`, `ExcelSheet`, `ExcelCell`).
- The converter bridges both domains: `ExcelToPdfConverter`.

---

## Class Design

### Sealed Classes

All public classes must be `sealed` to prevent unintended inheritance:

```csharp
public sealed class PdfDocument { }
public sealed class PdfPage { }
public sealed class PdfTextBlock { }
```

### Internal Constructors

Implementation types use `internal` constructors to prevent external instantiation while allowing internal factory usage:

```csharp
public sealed class PdfPage
{
    internal PdfPage(float width, float height) { }
}
```

### Immutable Collections

Expose collections as `IReadOnlyList<T>` backed by private `List<T>` fields:

```csharp
private readonly List<PdfPage> _pages = [];
public IReadOnlyList<PdfPage> Pages => _pages;
```

### Value Types

Use `readonly struct` with `IEquatable<T>` for small, immutable value types:

```csharp
public readonly struct PdfColor : IEquatable<PdfColor>
{
    public float R { get; }
    public float G { get; }
    public float B { get; }
}
```

Value types must:
- Override `Equals(object?)`, `GetHashCode()`, and `ToString()`
- Implement `operator ==` and `operator !=`
- Use `Math.Clamp()` to validate ranges in the constructor

---

## API Design Patterns

### Fluent Builder Pattern

Methods that modify state should return the same object to enable chaining:

```csharp
public PdfPage AddText(string text, float x, float y, float fontSize = 12, PdfColor? color = null)
{
    _textBlocks.Add(new PdfTextBlock(text, x, y, fontSize, color));
    return this;
}
```

Usage:
```csharp
page.AddText("Line 1", 50, 700)
    .AddText("Line 2", 50, 680);
```

### Factory Methods

Use static factory methods for type creation from external formats:

```csharp
public static PdfColor FromRgb(byte r, byte g, byte b)
    => new(r / 255f, g / 255f, b / 255f);

public static PdfColor FromHex(string hex) { /* ... */ }
```

### Named Static Presets

Provide commonly used values as static properties:

```csharp
public static PdfColor Black => new(0, 0, 0);
public static PdfColor Red => new(1, 0, 0);
```

### Default Parameters

Use default parameter values to keep the API simple while allowing customization:

```csharp
public PdfPage AddPage(float width = 612, float height = 792)
public PdfPage AddText(string text, float x, float y, float fontSize = 12, PdfColor? color = null)
```

Default page size is US Letter (612 × 792 points).

---

## PDF 1.4 Specification Conventions

MiniPdf generates PDF 1.4 compliant output. All contributors working on PDF generation must follow these conventions.

### File Structure

Every PDF file must contain:
1. **Header**: `%PDF-1.4` followed by a binary comment marker (`%\xe2\xe3\xcf\xd3`)
2. **Object tree**: Catalog → Pages → Page objects with content streams
3. **Cross-reference table** (`xref`): Maps object numbers to byte offsets
4. **Trailer**: Contains `/Size` and `/Root` references, followed by `startxref` and `%%EOF`

### Object Numbering

- Object 1: Catalog (`/Type /Catalog /Pages 2 0 R`)
- Object 2: Pages collection (`/Type /Pages /Kids [...] /Count N`)
- Object 3: Font resource (`/Type /Font /Subtype /Type1 /BaseFont /Helvetica`)
- Objects 4+: Page and content stream pairs

### Font Handling

- Use only the built-in Helvetica font with WinAnsiEncoding.
- Reference as `/F1` in content streams.
- Approximate character width: `fontSize * 0.5` points (for layout calculations).
- No font embedding is performed.

### Content Streams

Content streams use PDF text operators within `BT` / `ET` blocks:

| Operator | Purpose                        | Example               |
|----------|--------------------------------|-----------------------|
| `Tf`     | Set font and size              | `/F1 12 Tf`           |
| `Td`     | Move text position             | `50 700 Td`           |
| `Tj`     | Show text string               | `(Hello World) Tj`    |
| `rg`     | Set fill color (RGB)           | `1.000 0.000 0.000 rg`|

### Text Positioning

Each text block uses absolute positioning:
1. Move to position: `{x} {y} Td`
2. Render text: `({escaped_text}) Tj`
3. Reset to origin: `{-x} {-y} Td`

This reset-after-render pattern ensures each block is independently positioned.

### String Escaping

PDF strings inside parentheses must escape:
- `\` → `\\`
- `(` → `\(`
- `)` → `\)`
- `\r` → `\\r`
- `\n` → `\\n`

### Numeric Formatting

All floating-point values in PDF output must use `CultureInfo.InvariantCulture`:

```csharp
var x = block.X.ToString(CultureInfo.InvariantCulture);
```

Color values use `F3` format (three decimal places):
```csharp
var r = color.R.ToString("F3", CultureInfo.InvariantCulture);
```

### Encoding

All PDF output is written using `Encoding.ASCII`. Content stream bytes are calculated from ASCII encoding for the `/Length` field.

---

## Color Handling

### RGB Color Model

Colors use the RGB model with component values in the 0.0–1.0 range.

### Input Formats

The library accepts colors in these formats:
- **Float RGB** (0.0–1.0): `new PdfColor(1.0f, 0.0f, 0.0f)`
- **Byte RGB** (0–255): `PdfColor.FromRgb(255, 0, 0)`
- **Hex string** (6-char): `PdfColor.FromHex("FF0000")` or `PdfColor.FromHex("#FF0000")`
- **ARGB hex** (8-char): `PdfColor.FromHex("FFFF0000")` — alpha channel is ignored

### Clamping

Constructor values are clamped to valid range using `Math.Clamp(value, 0f, 1f)`.

### PDF Output

Colors are emitted using the `rg` operator. Black (0,0,0) is always emitted as `0 0 0 rg`. Non-black colors use three-decimal format: `1.000 0.000 0.000 rg`.

### Fallback

Invalid hex strings return `PdfColor.Black`. Null or empty hex strings return `PdfColor.Black`.

---

## Excel Integration

### Zero-Dependency Approach

Excel (.xlsx) files are parsed using only built-in .NET APIs:
- `System.IO.Compression.ZipArchive` for the .xlsx container
- `System.Xml.Linq.XDocument` for XML content

### Parsed Components

The reader extracts from the .xlsx archive:
- `xl/sharedStrings.xml` — shared string table
- `xl/styles.xml` — font colors (ARGB hex, indexed colors)
- `xl/workbook.xml` — sheet names and ordering
- `xl/worksheets/sheet{N}.xml` — cell data

### Color Resolution Chain

Cell colors are resolved through a chain:
1. Cell `s` attribute → style index
2. Style index → `cellXfs` entry → `fontId`
3. `fontId` → font entry → `color` element
4. Color element: `rgb` attribute (ARGB hex) → `indexed` attribute → fallback to null

### Column Width Calculation

- Natural width = `max(cellTextLength, 3) * fontSize * 0.5`
- If total natural width exceeds usable page width, columns are proportionally scaled down
- Column padding is subtracted before scaling

### Pagination

- Rows that would exceed the bottom margin trigger a new page
- Each sheet starts on a new page
- Empty sheets are skipped (at least one empty page is created if no sheets have data)

---

## Error Handling and Validation

### Defensive Defaults

- Use null-coalescing (`??`) to provide sensible defaults:
  ```csharp
  options ??= new ConversionOptions();
  Color = color ?? PdfColor.Black;
  ```

- Use early returns for empty or invalid input:
  ```csharp
  if (string.IsNullOrEmpty(text)) return this;
  if (sheet.Rows.Count == 0) return;
  ```

### Value Clamping

Use `Math.Clamp()` instead of throwing exceptions for out-of-range values:
```csharp
R = Math.Clamp(r, 0f, 1f);
```

### Graceful Fallback

- Invalid hex colors return `PdfColor.Black` instead of throwing
- Missing Excel sheets are skipped
- Missing shared strings or styles produce empty/default values

### No Exceptions for Invalid Input

The library prefers returning default values over throwing exceptions for malformed input. This ensures robustness when processing real-world Excel files that may have unexpected formats.

---

## Documentation

### XML Documentation Comments

All public types and members must have XML documentation comments (`///`):

```csharp
/// <summary>
/// Adds a text block at the specified position.
/// </summary>
/// <param name="text">The text to render.</param>
/// <param name="x">X position in points from the left edge.</param>
/// <param name="y">Y position in points from the bottom edge.</param>
/// <param name="fontSize">Font size in points (default: 12).</param>
/// <param name="color">Text color (default: black).</param>
/// <returns>The current page for chaining.</returns>
public PdfPage AddText(string text, float x, float y, float fontSize = 12, PdfColor? color = null)
```

### Required Tags

- `<summary>` on all public types and members
- `<param>` on all public method parameters
- `<returns>` on all public methods with return values
- `<inheritdoc />` when overriding standard methods (`Equals`, `GetHashCode`, `ToString`)

### Units

Always specify units in documentation: "in points", "in points from the left edge", "0.0–1.0 range".

---

## Testing

### Framework

- Use **xUnit** as the test framework
- Use **Coverlet** for code coverage

### Test Class Organization

Each public class should have a corresponding test class:
- `PdfDocument` → `PdfDocumentTests`
- `PdfColor` → `PdfColorTests`
- `ExcelToPdfConverter` → `ExcelToPdfConverterTests`

### Test Naming Convention

Test method names follow the pattern:

```
{MethodUnderTest}_{Scenario}_{ExpectedBehavior}
```

Examples:
```csharp
AddPage_DefaultSize_CreatesUsLetterPage()
FromHex_8CharArgb_SkipsAlpha()
Convert_ManyRows_CreatesMultiplePages()
Save_EscapesSpecialCharacters()
```

### Test Structure

Each test should:
1. **Arrange** — set up the test data
2. **Act** — call the method under test
3. **Assert** — verify the result

### PDF Output Validation

Tests validate PDF output by:
- Converting to byte array with `doc.ToArray()`
- Decoding to ASCII string
- Asserting on structural markers (`%PDF-1.4`, `%%EOF`, `/Type /Page`)
- Asserting on content presence (text, color operators, font references)

### Temporary Files

Tests that create files must:
- Use `Path.GetTempPath()` with unique names (`Guid.NewGuid()`)
- Clean up in a `finally` block

### In-Memory Excel Files

Tests create xlsx files in-memory using `ZipArchive` with `MemoryStream`. Helper methods (`CreateSimpleExcel`, `CreateColoredExcel`) build minimal valid xlsx archives for testing.

---

## Security

### .NET Security Policy

- **Nullable reference types** are enabled (`<Nullable>enable</Nullable>`) to prevent null reference exceptions.
- **Input validation** is performed on all public API boundaries.
- **No external dependencies** in the library — minimizes the attack surface.
- **Path traversal protection**: File paths are passed through to `System.IO.File` methods which handle OS-level path validation.
- **XML parsing**: `XDocument.Load()` is used with default settings. When processing untrusted xlsx files, consider the risk of XML-related attacks.
- **Stream handling**: All streams are properly disposed using `using` statements.
- **No user-controlled format strings**: All `ToString()` calls use explicit format specifiers and `CultureInfo.InvariantCulture`.

---

## Dependencies

### Zero-Dependency Policy

The MiniPdf library must have **zero external NuGet dependencies**. Only built-in .NET BCL types are allowed:
- `System.IO.Compression` for zip/xlsx handling
- `System.Xml.Linq` for XML parsing
- `System.Text` for encoding
- `System.Globalization` for culture-invariant formatting

### Test Dependencies

Test projects may use:
- `xunit` — test framework
- `xunit.runner.visualstudio` — test runner
- `coverlet.collector` — code coverage
- `Microsoft.NET.Test.Sdk` — test infrastructure

### Adding Dependencies

Before adding any new dependency:
1. Verify the functionality cannot be achieved with built-in .NET APIs
2. Check for security advisories on the package
3. Ensure it is compatible with the target framework (net9.0)
4. The test project allows NuGet dependencies; the library project does not
