[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [string]$Sheet = "Scan Data",

    [Parameter(Mandatory = $false)]
    [string]$SplitColumn = "BU"
)

try {
    # Check if file exists and is readable
    if (-not (Test-Path $FilePath)) {
        throw "Error: File '$FilePath' does not exist."
    }

    # Verify it's an Excel file
    $extension = [System.IO.Path]::GetExtension($FilePath)
    if ($extension -ne ".xlsx" -and $extension -ne ".xls") {
        throw "Error: File '$FilePath' is not an Excel file (.xlsx or .xls)."
    }

    Write-Host "Reading file: $FilePath" -ForegroundColor Cyan

    # Import the Excel file - this requires the ImportExcel module
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "Error: ImportExcel module is not installed. Install it with: Install-Module -Name ImportExcel -Scope CurrentUser"
    }

    # Import the specified sheet
    try {
        $data = Import-Excel -Path $FilePath -WorksheetName $Sheet -ErrorAction Stop
    }
    catch {
        throw "Error: Sheet '$Sheet' not found in file '$FilePath'. Please verify the sheet name."
    }

    if (-not $data -or $data.Count -eq 0) {
        throw "Error: Sheet '$Sheet' is empty or contains no data."
    }

    Write-Host "Successfully read sheet '$Sheet' with $($data.Count) rows" -ForegroundColor Green

    # Check if the split column exists
    $firstRow = $data[0]
    $columnExists = $firstRow.PSObject.Properties.Name -contains $SplitColumn

    if (-not $columnExists) {
        $availableColumns = $firstRow.PSObject.Properties.Name -join ", "
        throw "Error: Column '$SplitColumn' not found in sheet '$Sheet'. Available columns: $availableColumns"
    }

    Write-Host "Column '$SplitColumn' found successfully" -ForegroundColor Green

    # Get unique values in the split column
    $uniqueValues = $data | Select-Object -ExpandProperty $SplitColumn -Unique | Where-Object { $_ -ne $null -and $_ -ne "" }
    
    Write-Host "`nFound $($uniqueValues.Count) unique value(s) in column '$SplitColumn':" -ForegroundColor Cyan
    $uniqueValues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

    # Get the base filename and directory
    $inputFileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($FilePath)

    Write-Host "`nSplitting file into separate Excel files..." -ForegroundColor Cyan

    # Process each unique value
    foreach ($value in $uniqueValues) {
        # Filter data for this value
        $filteredData = $data | Where-Object { $_.$SplitColumn -eq $value }
        
        # Create output filename: {value}-{originalfilename}.xlsx
        # Sanitize the value to remove invalid filename characters
        $sanitizedValue = $value -replace '[\\/:*?"<>|]', '_'
        $outputFileName = "${sanitizedValue}-${inputFileName}.xlsx"
        $outputPath = Join-Path $outputDirectory $outputFileName

        Write-Host "  Creating: $outputFileName ($($filteredData.Count) rows)" -ForegroundColor Gray

        # Export to Excel with frozen header and auto-filter
        $filteredData | Export-Excel -Path $outputPath -WorksheetName $Sheet -AutoFilter -FreezeTopRow -BoldTopRow -AutoSize

        Write-Host "    Saved: $outputPath" -ForegroundColor Green
    }

    Write-Host "`nSuccessfully split into $($uniqueValues.Count) file(s)!" -ForegroundColor Green

}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
