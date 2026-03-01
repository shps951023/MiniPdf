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

// File to file
MiniPdf.ConvertToPdf("data.xlsx", "data.pdf");

// File to byte array
byte[] pdfBytes = MiniPdf.ConvertToPdf("data.xlsx");

// Stream to byte array
using var stream = File.OpenRead("data.xlsx");
byte[] pdfBytes = MiniPdf.ConvertToPdf(stream);
```

## Benchmark

MiniPdf output is compared against LibreOffice as the reference renderer across 30 classic test cases.

| Category | Count | Threshold |
|---|---|---|
| 🟢 Excellent | 25 | ≥ 95% |
| 🟡 Acceptable | 3 | 85% – 95% |
| 🔴 Needs Improvement | 2 | < 85% |

**Average overall score: ~95.4%** (text similarity + visual similarity combined)

### Visual Comparison

All 30 test cases comparing MiniPdf output (left) vs LibreOffice reference (right). Page 1 shown for multi-page results.

<table>
<tr><th>Test Case</th><th>MiniPdf</th><th>LibreOffice (Reference)</th><th>Score</th></tr>
<tr>
  <td><b>classic01</b><br/>Basic table with headers</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic01_basic_table_with_headers_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic01_basic_table_with_headers_p1_reference.png" width="320"/></td>
  <td>🟢 97.7%</td>
</tr>
<tr>
  <td><b>classic02</b><br/>Multiple worksheets</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic02_multiple_worksheets_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic02_multiple_worksheets_p1_reference.png" width="320"/></td>
  <td>🟢 99.3%</td>
</tr>
<tr>
  <td><b>classic03</b><br/>Empty workbook</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic03_empty_workbook_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic03_empty_workbook_p1_reference.png" width="320"/></td>
  <td>🟢 100.0%</td>
</tr>
<tr>
  <td><b>classic04</b><br/>Single cell</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic04_single_cell_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic04_single_cell_p1_reference.png" width="320"/></td>
  <td>🟢 100.0%</td>
</tr>
<tr>
  <td><b>classic05</b><br/>Wide table</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic05_wide_table_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic05_wide_table_p1_reference.png" width="320"/></td>
  <td>🟢 99.7%</td>
</tr>
<tr>
  <td><b>classic06</b><br/>Tall table</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic06_tall_table_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic06_tall_table_p1_reference.png" width="320"/></td>
  <td>🟢 95.1%</td>
</tr>
<tr>
  <td><b>classic07</b><br/>Numbers only</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic07_numbers_only_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic07_numbers_only_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic08</b><br/>Mixed text and numbers</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic08_mixed_text_and_numbers_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic08_mixed_text_and_numbers_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic09</b><br/>Long text<br/><i>MiniPdf: 1 page<br/>Reference: 12 pages</i></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic09_long_text_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic09_long_text_p1_reference.png" width="320"/></td>
  <td>🔴 20.3%</td>
</tr>
<tr>
  <td><b>classic10</b><br/>Special XML characters</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic10_special_xml_characters_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic10_special_xml_characters_p1_reference.png" width="320"/></td>
  <td>🟢 99.8%</td>
</tr>
<tr>
  <td><b>classic11</b><br/>Sparse rows</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic11_sparse_rows_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic11_sparse_rows_p1_reference.png" width="320"/></td>
  <td>🟢 100.0%</td>
</tr>
<tr>
  <td><b>classic12</b><br/>Sparse columns</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic12_sparse_columns_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic12_sparse_columns_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic13</b><br/>Date strings</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic13_date_strings_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic13_date_strings_p1_reference.png" width="320"/></td>
  <td>🟢 96.8%</td>
</tr>
<tr>
  <td><b>classic14</b><br/>Decimal numbers</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic14_decimal_numbers_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic14_decimal_numbers_p1_reference.png" width="320"/></td>
  <td>🟢 99.3%</td>
</tr>
<tr>
  <td><b>classic15</b><br/>Negative numbers</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic15_negative_numbers_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic15_negative_numbers_p1_reference.png" width="320"/></td>
  <td>🟢 99.8%</td>
</tr>
<tr>
  <td><b>classic16</b><br/>Percentage strings</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic16_percentage_strings_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic16_percentage_strings_p1_reference.png" width="320"/></td>
  <td>🟢 99.1%</td>
</tr>
<tr>
  <td><b>classic17</b><br/>Currency strings</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic17_currency_strings_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic17_currency_strings_p1_reference.png" width="320"/></td>
  <td>🟢 97.8%</td>
</tr>
<tr>
  <td><b>classic18</b><br/>Large dataset</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic18_large_dataset_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic18_large_dataset_p1_reference.png" width="320"/></td>
  <td>🟡 91.5%</td>
</tr>
<tr>
  <td><b>classic19</b><br/>Single column list</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic19_single_column_list_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic19_single_column_list_p1_reference.png" width="320"/></td>
  <td>🟢 99.7%</td>
</tr>
<tr>
  <td><b>classic20</b><br/>All empty cells</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic20_all_empty_cells_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic20_all_empty_cells_p1_reference.png" width="320"/></td>
  <td>🟢 100.0%</td>
</tr>
<tr>
  <td><b>classic21</b><br/>Header only</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic21_header_only_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic21_header_only_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic22</b><br/>Long sheet name</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic22_long_sheet_name_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic22_long_sheet_name_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic23</b><br/>Unicode text</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic23_unicode_text_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic23_unicode_text_p1_reference.png" width="320"/></td>
  <td>🟡 86.4%</td>
</tr>
<tr>
  <td><b>classic24</b><br/>Red text</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic24_red_text_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic24_red_text_p1_reference.png" width="320"/></td>
  <td>🟢 99.4%</td>
</tr>
<tr>
  <td><b>classic25</b><br/>Multiple colors</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic25_multiple_colors_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic25_multiple_colors_p1_reference.png" width="320"/></td>
  <td>🟢 99.3%</td>
</tr>
<tr>
  <td><b>classic26</b><br/>Inline strings</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic26_inline_strings_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic26_inline_strings_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic27</b><br/>Single row</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic27_single_row_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic27_single_row_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
<tr>
  <td><b>classic28</b><br/>Duplicate values</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic28_duplicate_values_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic28_duplicate_values_p1_reference.png" width="320"/></td>
  <td>🟢 99.8%</td>
</tr>
<tr>
  <td><b>classic29</b><br/>Formula results</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic29_formula_results_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic29_formula_results_p1_reference.png" width="320"/></td>
  <td>🟡 86.8%</td>
</tr>
<tr>
  <td><b>classic30</b><br/>Mixed empty and filled sheets</td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic30_mixed_empty_and_filled_sheets_p1_minipdf.png" width="320"/></td>
  <td><img src="tests/MiniPdf.Benchmark/reports/images/classic30_mixed_empty_and_filled_sheets_p1_reference.png" width="320"/></td>
  <td>🟢 99.9%</td>
</tr>
</table>

## License

MIT
