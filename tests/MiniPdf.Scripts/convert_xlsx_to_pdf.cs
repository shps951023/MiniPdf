#:project ../../src/MiniPdf/MiniPdf.csproj

using Mp = MiniPdf.MiniPdf;

// Resolve directories relative to this script file
var scriptDir = Path.GetDirectoryName(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar));
// When run via `dotnet run`, CWD is more reliable
var baseDir = Directory.GetCurrentDirectory();

var xlsxDir = args.Length > 0
    ? Path.GetFullPath(args[0])
    : Path.Combine(baseDir, "output");

var pdfDir = args.Length > 1
    ? Path.GetFullPath(args[1])
    : Path.Combine(baseDir, "pdf_output");

Directory.CreateDirectory(pdfDir);

var xlsxFiles = Directory.GetFiles(xlsxDir, "*.xlsx")
                         .OrderBy(f => f)
                         .ToArray();

if (xlsxFiles.Length == 0)
{
    Console.WriteLine($"No .xlsx files found in: {xlsxDir}");
    return 1;
}

Console.WriteLine($"Converting {xlsxFiles.Length} .xlsx files to PDF...");
Console.WriteLine($"  Input : {xlsxDir}");
Console.WriteLine($"  Output: {pdfDir}");
Console.WriteLine();

var passed = 0;
var failed = 0;

foreach (var xlsxPath in xlsxFiles)
{
    var name = Path.GetFileNameWithoutExtension(xlsxPath);
    var pdfPath = Path.Combine(pdfDir, name + ".pdf");

    try
    {
        Mp.ConvertToPdf(xlsxPath, pdfPath);
        var pdfSize = new FileInfo(pdfPath).Length;
        Console.WriteLine($"  OK  {name}.pdf ({pdfSize / 1024.0:F1} KB)");
        passed++;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ERR {name}: {ex.Message}");
        failed++;
    }
}

Console.WriteLine();
Console.WriteLine($"Done! Passed: {passed}, Failed: {failed}, Total: {xlsxFiles.Length}");

// Open output folder for manual inspection
if (OperatingSystem.IsWindows())
{
    Console.WriteLine($"\nOpening output folder...");
    System.Diagnostics.Process.Start("explorer.exe", pdfDir);
}

return failed > 0 ? 1 : 0;
