<#
.SYNOPSIS
    One-click benchmark: generate Excel → convert to PDF (MiniPdf + LibreOffice) → compare → report.

.DESCRIPTION
    This script orchestrates the full MiniPdf self-evolution pipeline on Windows.
    It installs Python dependencies, runs all steps, and opens the final report.

.EXAMPLE
    .\Run-Benchmark.ps1
    .\Run-Benchmark.ps1 -CompareOnly
    .\Run-Benchmark.ps1 -SkipReference
#>

param(
    [switch]$CompareOnly,
    [switch]$SkipGenerate,
    [switch]$SkipMiniPdf,
    [switch]$SkipReference,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Continue"
$ScriptRoot = $PSScriptRoot
$BenchmarkDir = Join-Path $ScriptRoot "tests" "MiniPdf.Benchmark"

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  MiniPdf Self-Evolution Benchmark Pipeline" -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan

# Step 0: Install Python dependencies
if (-not $SkipInstall) {
    Write-Host "[Step 0] Installing Python dependencies..." -ForegroundColor Yellow
    pip install openpyxl pymupdf --quiet 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  WARNING: pip install had issues. Continuing anyway..." -ForegroundColor DarkYellow
    } else {
        Write-Host "  OK" -ForegroundColor Green
    }
}

# Build args for Python pipeline
$pyArgs = @()
if ($CompareOnly) { $pyArgs += "--compare-only" }
if ($SkipGenerate) { $pyArgs += "--skip-generate" }
if ($SkipMiniPdf) { $pyArgs += "--skip-minipdf" }
if ($SkipReference) { $pyArgs += "--skip-reference" }

# Run the benchmark pipeline
Write-Host "`n[Running] python run_benchmark.py $($pyArgs -join ' ')`n" -ForegroundColor Yellow
Push-Location $BenchmarkDir
try {
    python run_benchmark.py @pyArgs
} finally {
    Pop-Location
}

# Open the report if it exists
$reportPath = Join-Path $BenchmarkDir "reports" "comparison_report.md"
if (Test-Path $reportPath) {
    Write-Host "`n[Done] Report: $reportPath" -ForegroundColor Green
    Write-Host "Opening report..." -ForegroundColor Cyan
    # Open in VS Code if available, otherwise notepad
    $code = Get-Command code -ErrorAction SilentlyContinue
    if ($code) {
        code $reportPath
    } else {
        Start-Process notepad.exe -ArgumentList $reportPath
    }
} else {
    Write-Host "`nNo report generated. Check the output above for errors." -ForegroundColor Red
}
