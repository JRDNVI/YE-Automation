# ==============================================================
# PowerShell Script to Dynamically Load and Run VBA Macros
# Creates a temporary workbook each run
# ==============================================================

# === CONFIG ===
$srcPath   = "C:\Users\coadyj\projects\YE-Automation\src"                
$tempDir   = "C:\Users\coadyj\projects\YE-Automation\temp"            
$macroName = "ImportSupervisorSheets"                        # Entry macro name

# === Ensure temp folder exists ===
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

# === Generate unique temp workbook path ===
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$tempFile   = Join-Path $tempDir ("TempHost_" + $timestamp + ".xlsm")

Write-Host "Creating temporary Excel host workbook: $tempFile"

# === Create Excel COM object ===
$excel = New-Object -ComObject Excel.Application
# Force-load VBIDE type library
[void][System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$excel.Visible = $false
$excel.DisplayAlerts = $false

# === Create blank workbook and save as .xlsm ===
$workbook = $excel.Workbooks.Add()
$workbook.SaveAs($tempFile, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)

# === Import all .bas and .cls files ===
Start-Sleep -Seconds 3
$vbaProject = $workbook.VBProject
$sourceFiles = Get-ChildItem -Path $srcPath -Recurse -Include *.bas, *.cls

foreach ($file in $sourceFiles) {
    Write-Host "Importing $($file.Name)"
    $vbaProject.VBComponents.Import($file.FullName)
}

# === Run the specified macro ===
try {
    Write-Host "Running macro: $macroName"
    $excel.Run($macroName)
    Write-Host "Macro executed successfully."
} catch {
    Write-Host "Error running macro: $($_.Exception.Message)"
}

# === Save, close, and cleanup ===
$workbook.Close($false)
$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# === Optional: Delete temporary file ===
try {
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force
        Write-Host "Deleted temp file: $tempFile"
    }
} catch {
    Write-Host "Could not delete temp file (it might be locked)."
}

Write-Host "Supervisor import completed from dynamic temp workbook."
