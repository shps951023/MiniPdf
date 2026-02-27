namespace MiniPdf;

/// <summary>
/// Main entry point for MiniPdf operations.
/// Provides simple methods for converting files to PDF format.
/// </summary>
public static class MiniPdf
{
    /// <summary>
    /// Converts an Excel (.xlsx) file to a PDF file.
    /// </summary>
    /// <param name="inputPath">Path to the source .xlsx file.</param>
    /// <param name="outputPath">Path for the output .pdf file.</param>
    public static void ConvertToPdf(string inputPath, string outputPath)
    {
        ExcelToPdfConverter.ConvertToFile(inputPath, outputPath);
    }

    /// <summary>
    /// Converts an Excel (.xlsx) file to a PDF byte array.
    /// </summary>
    /// <param name="inputPath">Path to the source .xlsx file.</param>
    /// <returns>A byte array containing the PDF data.</returns>
    public static byte[] ConvertToPdf(string inputPath)
    {
        var doc = ExcelToPdfConverter.Convert(inputPath);
        return doc.ToArray();
    }

    /// <summary>
    /// Converts an Excel (.xlsx) stream to a PDF byte array.
    /// </summary>
    /// <param name="inputStream">Stream containing .xlsx data.</param>
    /// <returns>A byte array containing the PDF data.</returns>
    public static byte[] ConvertToPdf(Stream inputStream)
    {
        var doc = ExcelToPdfConverter.Convert(inputStream);
        return doc.ToArray();
    }
}
