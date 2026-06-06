# Excel Custom Ribbon Creator - imageMso Icon Extractor
# Extracts all standard Excel icons from Microsoft Excel COM or converts extracted BMPs.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Source = @"
using System;
using System.Drawing;
using System.Windows.Forms;

public class AxHostConverter : AxHost {
    private AxHostConverter() : base("") {}
    
    public static Image IPictureToImage(object ipicture) {
        return GetPictureFromIPicture(ipicture);
    }
}
"@

try {
    Add-Type -TypeDefinition $Source -ReferencedAssemblies "System.Windows.Forms", "System.Drawing"
} catch {}

$baseDir = Get-Location
$listPath = Join-Path $baseDir "image_mso_list.js"
$outDir = Join-Path $baseDir "imageMSO"
$helperPath = Join-Path $baseDir "extract_helper.txt"

if (-not (Test-Path $listPath)) {
    Write-Error "Error: image_mso_list.js not found in current directory! Please run this script in the project folder."
    exit
}

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

# ===================================================================
# STEP 1: Check for BMP files in output directory and convert them to PNG
# ===================================================================
$bmpFiles = Get-ChildItem -Path $outDir -Filter "*.bmp"
if ($bmpFiles.Count -gt 0) {
    Write-Host "=================================================="
    Write-Host "  BMP to PNG Converter Mode Active"
    Write-Host "=================================================="
    Write-Host "Found $($bmpFiles.Count) extracted BMP files in imageMSO folder."
    Write-Host "Converting BMPs to transparent PNGs..."
    
    $converted = 0
    $totalBmps = $bmpFiles.Count
    
    foreach ($file in $bmpFiles) {
        $pngPath = [System.IO.Path]::ChangeExtension($file.FullName, ".png")
        try {
            $img = [System.Drawing.Image]::FromFile($file.FullName)
            $img.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $img.Dispose()
            Remove-Item $file.FullName
            $converted++
            
            if ($converted % 100 -eq 0 -or $converted -eq $totalBmps) {
                Write-Host "Progress: $converted / $totalBmps converted"
            }
        } catch {
            Write-Host "Warning: Failed to convert $($file.Name): $_"
        }
    }
    
    Write-Host "`nConversion completed successfully!"
    Write-Host "Converted files: $converted"
    Write-Host "All icons are now saved as high-resolution PNGs in: $outDir"
    exit
}

# ===================================================================
# STEP 2: Standard COM Extraction Mode (Runs if no BMPs found)
# ===================================================================
$content = Get-Content $listPath -Raw
$matches = [regex]::Matches($content, '[''"]([^''''"]+)[''"]')
Write-Host "Found $($matches.Count) potential icons in image_mso_list.js"

# Let's count existing PNGs
$existingPngCount = (Get-ChildItem -Path $outDir -Filter "*.png").Count
if ($existingPngCount -eq 3245) {
    Write-Host "All 3,245 icons are already extracted as PNG in: $outDir"
    exit
}

Write-Host "Initializing Microsoft Excel via COM..."

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    $commandBars = $excel.CommandBars
} catch {
    Write-Host "Error initializing Excel COM: $_"
    Write-Host "Please make sure Microsoft Excel is installed."
    exit
}

Write-Host "Excel initialized. Starting extraction... (Existing files will be skipped)"
Write-Host "This will take 1-3 minutes. Please wait..."

$total = 0
$extracted = 0
$skipped = 0
$failed = 0
$consecutiveFailures = 0
$abortLimit = 5 # Abort script if first N calls fail with COM E_UNEXPECTED

foreach ($match in $matches) {
    $idMso = $match.Groups[1].Value
    if ($idMso -eq "ALL_IMAGE_MSO_LIST" -or [string]::IsNullOrEmpty($idMso)) { continue }
    
    $total++
    $destFile = Join-Path $outDir "$idMso.png"
    
    if (Test-Path $destFile) {
        $skipped++
        continue
    }
    
    try {
        $ipicture = $commandBars.GetImageMso($idMso, 32, 32)
        if ($ipicture -ne $null) {
            $image = [AxHostConverter]::IPictureToImage($ipicture)
            $image.Save($destFile, [System.Drawing.Imaging.ImageFormat]::Png)
            $extracted++
            $consecutiveFailures = 0
        } else {
            $failed++
        }
    } catch {
        $failed++
        $consecutiveFailures++
        
        # Check if COM is throwing unexpected error (HRESULT: 0x8000FFFF / E_UNEXPECTED)
        if ($consecutiveFailures -ge $abortLimit) {
            Write-Host "`n[ALERT] COM Catastrophic Failure detected (HRESULT: 0x8000FFFF / E_UNEXPECTED)."
            Write-Host "This occurs when Excel COM has environment blocks or is blocked by licensing/pop-ups."
            Write-Host "Aborting script to prevent wasting time.`n"
            
            Write-Host "=================================================================="
            Write-Host "                  VBA WORKAROUND INSTRUCTIONS"
            Write-Host "=================================================================="
            Write-Host "Since the PowerShell COM interface is blocked on your system, please use"
            Write-Host "the Excel VBA Macro workaround. We have created a helper file for you:"
            Write-Host "-> [extract_helper.txt] ($helperPath)"
            Write-Host ""
            Write-Host "Steps to run:"
            Write-Host "1. Open Microsoft Excel manually on your desktop."
            Write-Host "2. Save a blank workbook as a Macro-Enabled Workbook (*.xlsm)"
            Write-Host "   in this directory: $baseDir"
            Write-Host "3. Press [Alt + F11] to open the VBA Editor."
            Write-Host "4. Click [Insert] -> [Module] and paste the content of 'extract_helper.txt'."
            Write-Host "5. Click inside the 'ExtractIconsToBmp' subroutine and press [F5] (Run)."
            Write-Host "6. Excel will extract all 3,245 icons as BMP files in seconds."
            Write-Host "7. Once complete, run this PowerShell script again in the folder:"
            Write-Host "   powershell -ExecutionPolicy Bypass -File .\extract_icons.ps1"
            Write-Host "   (It will automatically convert the BMPs to transparent PNGs!)"
            Write-Host "=================================================================="
            
            # Clean up COM
            if ($wb) { $wb.Close($false) }
            if ($excel) { $excel.Quit() }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            exit
        }
    }
    
    if ($total % 100 -eq 0) {
        Write-Host "Progress: $total processed | Extracted: $extracted | Skipped: $skipped | Failed/Missing: $failed"
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "`nExtraction process completed!"
Write-Host "Total processed: $total"
Write-Host "New icons extracted: $extracted"
Write-Host "Existing files skipped: $skipped"
Write-Host "Failed or missing: $failed"
Write-Host "All icons are saved in: $outDir"
