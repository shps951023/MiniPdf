# MiniPdf Roadmap

This document tracks planned features and improvements for the MiniPdf library.

## v0.1.0 (Current)

- [x] Text-to-PDF with positioned text
- [x] Auto-wrapped text within width bounds
- [x] Excel-to-PDF with automatic column layout
- [x] Multi-sheet support
- [x] Automatic pagination
- [x] Text color support (PdfColor)
- [x] Excel font color reading (ARGB hex, indexed colors)
- [x] PDF metadata (Title, Author, Subject, Keywords, Creator)
- [x] Input validation on all public APIs
- [x] NuGet packaging with README

## v0.2.0 (Planned)

- [ ] Bold / italic / underline text (Helvetica-Bold, Helvetica-Oblique)
- [ ] Per-cell font size from Excel
- [ ] Text alignment (left, center, right)
- [ ] Cell background / fill color
- [ ] Column width from Excel (`<col>` widths)
- [ ] Row height from Excel

## v0.3.0 (Planned)

- [ ] Merged cell support
- [ ] Number and date formatting from Excel format codes
- [ ] Hyperlink annotations
- [ ] Headers and footers (page numbers, sheet name, date)
- [ ] Page setup from Excel (orientation, paper size, margins)

## Future Considerations

- [ ] Image support (JPEG/PNG embedding)
- [ ] Multi-target frameworks (netstandard2.0, net6.0, net8.0)
- [ ] Async conversion APIs
- [ ] Configuration class for advanced options
- [ ] Streaming reads/writes for large file optimization
- [ ] PDF security features (passwords, permissions)
