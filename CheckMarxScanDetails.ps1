# This script retrieves Checkmarx scan details and statistics for all projects, exports them to a CSV file, and creates an Excel report with pivot tables.
# It requires the ImportExcel module for Excel operations.
# Ensure you have the ImportExcel module installed: Install-Module -Name ImportExcel -Scope CurrentUser
# Set debug preference
# $VerbosePreference = "Continue"
$DebugPreference = "SilentlyContinue"
$startTime = Get-Date
# Bypass SSL certificate validation (use with caution, remove if certificate is trusted)
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

# which OS
if ($PSVersionTable.Platform -eq "Unix" -and $PSVersionTable.OS -like "*Darwin*") {
    $mac = $true
} elseif ($PSVersionTable.Platform -eq "Win32NT") {
    $win = $true
} else {
    Write-Output "Can't determine OS from $PSVersionTable.Platform, exiting"
    exit 1
}

# Define URLs
$baseUrl = "https://checkmarxweb.corp.allscripts.com/cxrestapi/"
$tokenUrl = "$($baseURL)auth/identity/connect/token"
$projectsUrl = "$($baseURL)projects"
$scansUrlBase = "$($baseURL)sast/scans"
$statsUrlBase = "$($baseURL)sast/scans"
$teamsUrl = "$($baseURL)auth/Teams"

# Output file paths
$today = Get-Date
$fileDate = $today.ToString("MM-dd-yyyy")
if ($mac) {
    $baseDir = "/Users/$env:USER/powershell/"
    $outputCsv = "$($baseDir)CheckmarxScanStats.csv"
    $excelPath = "$($baseDir)CheckmarxScanStats for $($fileDate).xlsx"
    $pwPath = "$($baseDir)APIKey.chm"
}
else {
    $baseDir = "C:\powershell\"
    $outputCsv = "$($baseDir)CheckmarxScanStats.csv"
    $excelPath = "$($baseDir)CheckmarxScanStats for $($fileDate).xlsx"
    $pwPath = "$($baseDir)APIKey.chm"
}
Write-Host $pwPath
# Define authentication parameters
$body = @{
    username = "username=SunriseAPIUser"
    password = "password=" + (Get-Content $pwPath)
#([System.Net.NetworkCredential]::new("", (Get-Content $pwPath | ConvertTo-SecureString)).Password)
    grant_type = "grant_type=password"
    scope = "scope=sast_rest_api"
    client_id = "client_id=resource_owner_client"
    client_secret = "client_secret=014DF517-39D1-4453-B7B3-9930C563627C"
}
$bodyString = $body.Values -join "&"

if (Test-Path $outputCsv) {
    Remove-Item $outputCsv
}
if (Test-Path $excelPath) {
    Remove-Item $excelPath
}

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/x-www-form-urlencoded")

try {
    # step 1: get auth token
    try {
            $response = Invoke-RestMethod $tokenUrl -Method 'POST' -Headers $headers -Body $bodyString -SkipCertificateCheck
            $accessToken = $response.access_token
        }
        catch {
            Write-Error "Failed to get access token. StatusCode: $($_.Exception.Response.StatusCode), Reason: $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $errorResponse = $reader.ReadToEnd()
                Write-Output "Debug: Server error response: $errorResponse"
            }
            throw
        }

    # Step 2: Get all teams to map teamId to team name
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Accept"        = "application/json"
    }
    $teamsResponse = Invoke-RestMethod -Uri $teamsUrl -Method Get -Headers $headers -SkipCertificateCheck
    $teamMapping = @{}
    write-debug "Found $($teamsResponse.Count) teams"
    foreach ($team in $teamsResponse) {
        $rawTeamName = $team.fullName
        Write-Debug "  Team ID: $($team.id), Raw fullName: $rawTeamName"
        $teamNameParts = $rawTeamName.Split('/')
        if ($teamNameParts.Count -eq 2) {
            $teamName = $teamNameParts[1]
        }
        else {
            $teamName = $teamNameParts[2]
        }
        $teamMapping[$team.id] = $teamName
    }

    # Step 3: Get all projects
    $projectsResponse = Invoke-RestMethod -Uri $projectsUrl -Method Get -Headers $headers -SkipCertificateCheck

    # Step 4: For each project, get the latest scan and its statistics
    Write-Debug "Retrieving scan statistics for each project, found $($projectsResponse.Count) projects"
    
    $results=@()
    foreach ($project in $projectsResponse) {
        Write-Debug "Processing Project ID: $($project.id), Name: $($project.name)"

        # Initialize default values
        $result = [PSCustomObject]@{
            "ProjectId"        = $project.id
            "ProjectName"      = $project.name
            "TeamName"         = $teamMapping[$project.teamId] 
            "ScanId"           = ""
            "ScanDate"         = ""
            "CriticalSeverity" = 0
            "HighSeverity"     = 0
            "MediumSeverity"   = 0
            "LowSeverity"      = 0
            "InfoSeverity"     = 0
            "LinesOfCode"      = 0
            "FileCount"        = 0
        }

        # Get the latest scans for the project
        $scansUrl = "$scansUrlBase`?projectId=$($project.id)&last=10&scanStatus=7"
        try {
            $scansResponse = Invoke-RestMethod -Uri $scansUrl -Method Get -Headers $headers -SkipCertificateCheck

            if ($scansResponse -and $scansResponse.Count -gt 0) {
                $latestScan = $scansResponse[0]
                $result.ScanId = $latestScan.id
                $result.ScanDate = $latestScan.dateAndTime.startedOn
                $result.LinesOfCode = $latestScan.scanState.linesOfCode
                $result.FileCount = $latestScan.scanState.filesCount

                # Get scan statistics
                $statsUrl = "$statsUrlBase/$($latestScan.id)/resultsStatistics"
                try {
                    $statsResponse = Invoke-RestMethod -Uri $statsUrl -Method Get -Headers $headers -SkipCertificateCheck

                    # Populate statistics (assuming API fields; critical not standard, set to 0)
                    $result.CriticalSeverity = $statsResponse.criticalSeverity
                    $result.HighSeverity = $statsResponse.highSeverity
                    $result.MediumSeverity = $statsResponse.mediumSeverity
                    $result.LowSeverity = $statsResponse.lowSeverity
                    $result.InfoSeverity = $statsResponse.infoSeverity
                    $result.LinesOfCode = $latestScan.scanState.linesOfCode
                    $result.FileCount = $latestScan.scanState.filesCount

                }
                catch {
                    Write-Warning "Failed to retrieve statistics for Scan ID: $($latestScan.id). Error: $_"
                }
            }
            else {
                Write-Output "No scans found for $($project.name) $($teamMapping[$project.teamId]) at $scansUrl."
            }
        }
        catch {
            Write-Warning "Failed to retrieve scans for Project ID: $($project.id). Error: $_"
        }

        # Add result to collection
        $results += $result
    }

    # Step 5: Export to CSV
    $results | Export-Csv -Path $outputCsv -NoTypeInformation -Delimiter ","
    Write-Debug "Data exported to $outputCsv"
}
catch {
    # Handle errors for authentication, teams, or projects request
    Write-Error "Error: $_"
    exit 1
}
finally {
    # Reset certificate validation callback
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
}

# Import required module
# Import-Module ImportExcel
Remove-Module ImportExcel -ErrorAction SilentlyContinue
Import-Module ImportExcel -Version 7.8.6

# Create a hashtable with column names and values (e.g., representing column widths)
$labels = New-Object System.Collections.Specialized.OrderedDictionary
$labels.Add("Project ID", 10)
$labels.Add("Project Name", 45)
$labels.Add("BU", 12)
$labels.Add("Scan ID", 12)
$labels.Add("Scan Date", 11)
$labels.Add("Critical", 10)
$labels.Add("High", 10)
$labels.Add("Medium", 10)
$labels.Add("Low", 10)
$labels.Add("Info", 10)
$labels.Add("LOC", 12)
$labels.Add("Files", 10)

# Read CSV and create Excel workbook
$data = Import-Csv $outputCsv
$data = $data | Sort-Object -Property "TeamName", "ProjectName"

# Calculate the totals for the specified column
$totalLOC = ($data | Measure-Object -Property "LinesOfCode" -Sum).Sum
$totalCritical = ($data | Measure-Object -Property "CriticalSeverity" -Sum).Sum
$totalHigh = ($data | Measure-Object -Property "HighSeverity" -Sum).Sum

# export to Excel
$excel = $data | Export-Excel -Path $excelPath -WorksheetName "Scan Data" -TableName "CsvData" -TableStyle Medium9 -AutoSize -PassThru 
$worksheet = $excel.Workbook.Worksheets["Scan Data"]

# Set number column formats
for ($col = 6; $col -le 12; $col++) {
    $worksheet.Column($col).Style.Numberformat.Format = "#,##0"
}

# set column widths and labels
$idx = 1
foreach ($n in $labels.Keys) {
    $worksheet.Column($idx).Width = $labels[$n]
    $worksheet.Cells[1, $idx].Value = $n
    $idx++
}

# Get the actual number of rows in the worksheet
$maxRows = $worksheet.Dimension.Rows

# Iterate through all cells in the Scan Date column (from row 2 to maxRows) to set date format
for ($row = 2; $row -le $maxRows; $row++) {
    $dateWithoutMs = $worksheet.Cells[$row, 5].Value -replace "\.\d+$", ""
    try {
        $worksheet.Cells[$row, 5].Value = [DateTime]::ParseExact($dateWithoutMs, "yyyy-MM-ddTHH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
        $worksheet.Cells[$row, 5].Style.Numberformat.Format = "mm/dd/yyyy"
    }
    catch
    {
    }
}

# Create or access destination worksheet for pivot table
$destWorksheetName = "Critical And High By BU"
$destWorksheet = $excel.Workbook.Worksheets.Add($destWorksheetName)

# Create pivot table
try {
    # Define the source data range
    $dataRange = $worksheet.Dimension
    $pivotTableRange = $worksheet.Cells[$dataRange.Address]
    
    # Add pivot table
    $pivotTable = $destWorksheet.PivotTables.Add($destWorksheet.Cells["A1"], $pivotTableRange, "CriticalByBU")
    $pivotTable.DataOnRows = $false  # Place data fields in columns, not rows

    # Add row field (BU)
    $rowField = $pivotTable.Fields["BU"]
    $pivotTable.RowFields.Add($rowField)
    
    # Add data field (Critical, summed)
    $dataField = $pivotTable.Fields["Critical"]
    $dataFieldItem = $pivotTable.DataFields.Add($dataField)
    $dataFieldItem.Function = [OfficeOpenXml.Table.PivotTable.DataFieldFunctions]::Sum
    $dataFieldItem.Name = "Sum of Critical"
    $dataFieldItem.Format = "#,##0"

    # Add data field (High, summed)
    $dataField2 = $pivotTable.Fields["High"]
    $dataFieldItem2 = $pivotTable.DataFields.Add($dataField2)
    $dataFieldItem2.Function = [OfficeOpenXml.Table.PivotTable.DataFieldFunctions]::Sum
    $dataFieldItem2.Name = "Sum of High"
    $dataFieldItem2.Format = "#,##0"

    # Enable grand totals for rows and columns
    $pivotTable.RowGrandTotals = $true
    } catch {
    Write-Error "Failed to create pivot table: $_.Exception.Message"
    Write-Error $_.ScriptStackTrace
}

# Add total row in data sheet
$rowCount = $worksheet.Dimension.Rows
$totalRow = $rowCount + 1
$worksheet.Cells[$totalRow, 1].Value = "Totals"
$worksheet.Cells[$totalRow, 6].Value = $totalCritical
$worksheet.Cells[$totalRow, 7].Value = $totalHigh
$worksheet.Cells[$totalRow, 11].Value = $totalLOC

# Optional: Apply bold formatting to the total row
$worksheet.Cells[$totalRow, 1].Style.Font.Bold = $true
$worksheet.Cells[$totalRow, 6].Style.Font.Bold = $true
$worksheet.Cells[$totalRow, 7].Style.Font.Bold = $true
$worksheet.Cells[$totalRow, 11].Style.Font.Bold = $true


# Save and close the workbook
$excel.Save()
$excel.Dispose()
$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host "`nSuccess! Duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
