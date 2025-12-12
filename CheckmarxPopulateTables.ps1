[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]
    [string]$GivePath_To_Generate_Report,

    [Parameter(Mandatory = $True)]
    [string]$CheckmarxUsername,

    [Parameter(Mandatory = $True)]
    [string]$CheckmarxPassword,

    [string]$whichBU   # choose which group of products to load
)
$startTime = Get-Date
<#
Checkmarx SAST stores multiple scans for each product/project.
This program populates the Artifacts table if there is no Artifact record,
then adds scan rows to the Scans table for each historical scan if not already present.
The newest scan for each product is included in the Excel report.
#>

# Build filename timestamp in MMM dd YY format
$timestamp = Get-Date -Format "MMM_dd_yy"
$scanTool = "Checkmarx"
$scanType = "SAST"
# Build filename: GroupNameCheckmarxSASTMMM dd yy.xlsx
$filename = "$whichBU" + "_$scanTool" + "_$scanType" + "_" + $timestamp + ".xlsx"

# Dot-source the utility function
. "$PSScriptRoot\Test-FileWritable.ps1"

# Test if output file is writable (may be open in Excel)
$fullPath = Join-Path $GivePath_To_Generate_Report $filename
if (Test-Path $fullPath) {
    if (-not (Test-FileWritable $fullPath)) {
        Write-Error "Cannot proceed: output file is locked or not writable."
        exit 1
    }
}

try {
    $results = @()
    $checkmarxBaseURL = "CheckmarxBaseURL"
    
    # Build paths for config files
    $jsonTemplatePath = Join-Path $PSScriptRoot "ProductGroups.json"
    $jsonLocalPath = Join-Path $PSScriptRoot "ProductGroups.local.json"

    # Try to use local config first (with real tokens), fall back to template
    if (Test-Path $jsonLocalPath) {
        $configPath = $jsonLocalPath
        Write-Host "Using local config: $jsonLocalPath" -ForegroundColor Gray
    }
    elseif (Test-Path $jsonTemplatePath) {
        $configPath = $jsonTemplatePath
        Write-Host "Using template config: $jsonTemplatePath (consider creating ProductGroups.local.json with real data)" -ForegroundColor Yellow
    }
    else {
        throw "Config files not found. Create either:$([Environment]::NewLine)  - $jsonTemplatePath (template)$([Environment]::NewLine)  - $jsonLocalPath (local with real data, gitignored)"
    }

    # Read JSON file containing all product groups
    $groups = Get-Content $configPath | ConvertFrom-Json

    # Select the group requested on the command line (case-sensitive match)
    $matchedProperty = $groups.PSObject.Properties | Where-Object { $_.Name -ceq $whichBU }
    
    if (-not $matchedProperty) {
        throw "Group '$whichBU' not found in $configPath (case-sensitive match). Available groups: $($groups.PSObject.Properties.Name -join ', ')"
    }

    # Get Checkmarx products for this BU
    $checkmarxProperty = $matchedProperty.Value.PSObject.Properties | Where-Object { $_.Name -eq "Checkmarx" }
    
    if (-not $checkmarxProperty) {
        throw "No Checkmarx products defined for BU '$whichBU' in $configPath"
    }

    # Convert to hash table
    $CheckmarxProducts = @{}
    $checkmarxProperty.Value.PSObject.Properties | ForEach-Object {
        $CheckmarxProducts[$_.Name] = $_.Value
    }

    Write-Verbose "Found $($CheckmarxProducts.Count) Checkmarx products for BU '$whichBU'"

    # Authenticate with Checkmarx
    $authBody = @{
        username = $CheckmarxUsername
        password = $CheckmarxPassword
        grant_type = "password"
        scope = "sast_rest_api"
        client_id = "resource_owner_client"
        client_secret = "014DF517-39D1-4453-B7B3-9930C563627C"
    }

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $authResponse = Invoke-RestMethod -Method POST -Uri "$checkmarxBaseURL/cxrestapi/auth/identity/connect/token" -Body $authBody -ContentType "application/x-www-form-urlencoded"
    $token = $authResponse.access_token
    $headers = @{
        Authorization = "Bearer $token"
        Accept = "application/json"
    }

    Write-Host "Authenticated with Checkmarx" -ForegroundColor Green

    # Iterate through each Checkmarx product
    $CheckmarxProducts.GetEnumerator() | ForEach-Object {
        $ProductName = $_.Key
        $ProjectId = $_.Value

        Write-Host "`nCollecting Data for $whichBU $ProductName (Project ID: $ProjectId)..." -ForegroundColor Green

        # Get all scans for this project
        $scansUrl = "$checkmarxBaseURL/cxrestapi/sast/scans?projectId=$ProjectId&scanStatus=Finished"
        $scansResponse = Invoke-RestMethod -Method GET -Uri $scansUrl -Headers $headers
        
        if (-not $scansResponse -or $scansResponse.Count -eq 0) {
            Write-Host "No finished scans found for $ProductName" -ForegroundColor Yellow
            return
        }

        Write-Verbose "Found $($scansResponse.Count) scans for $ProductName"

        # Database setup
        $dbPath = "C:\Users\W988276\AppData\Local\monthlyReportDatabase\monthlyReport.db"
        $sqlitePath = "C:\sqlite\sqlite3.exe"
        
        # Enable foreign key constraints
        & $sqlitePath $dbPath "PRAGMA foreign_keys = ON;"
        
        # Get or create Artifact ID
        $query = @"
SELECT ID FROM Artifacts 
WHERE CheckmarxProduct = '$($ProductName.Replace("'", "''"))' 
  AND BusinessUnit = '$($whichBU.Replace("'", "''"))';
"@
        $artifactID = & $sqlitePath $dbPath $query
        
        if (-not $artifactID) {
            # Create new Artifact record
            $insertArtifact = @"
PRAGMA foreign_keys = ON;
INSERT INTO Artifacts (BusinessUnit, CheckmarxProduct, SASTScans, RecentSAST) 
VALUES ('$($whichBU.Replace("'", "''"))', '$($ProductName.Replace("'", "''"))', 0, NULL);
SELECT last_insert_rowid();
"@
            $artifactID = & $sqlitePath $dbPath $insertArtifact
            Write-Verbose "Created new Artifact ID: $artifactID"
        }
        else {
            Write-Verbose "Found existing Artifact ID: $artifactID"
        }

        # Track the most recent scan for the Excel report
        $mostRecentScan = $null
        $mostRecentDate = [DateTime]::MinValue

        # Process each scan
        foreach ($scan in $scansResponse) {
            $scanId = $scan.id
            $scanDate = [DateTime]::Parse($scan.dateAndTime.finishedOn)
            $DateForScanTable = $scanDate.ToString("yyyy-MM-dd HH:mm:ss")
            $DateForReport = $scanDate.ToString("MMM dd, yyyy hh:mm:ss tt")

            Write-Verbose "Processing scan $scanId from $DateForScanTable"

            # Check if scan already exists
            $checkScan = @"
SELECT ID FROM Scans 
WHERE ArtifactID = $artifactID 
  AND ScanTool = '$scanTool' 
  AND ScanType = '$scanType' 
  AND ScanDateTime = '$DateForScanTable';
"@
            $existingScanID = & $sqlitePath $dbPath $checkScan
            
            if ($existingScanID) {
                Write-Verbose "Scan already exists (ID: $existingScanID) - skipping"
            }
            else {
                # Get scan statistics
                $statsUrl = "$checkmarxBaseURL/cxrestapi/sast/scans/$scanId/resultsStatistics"
                $stats = Invoke-RestMethod -Method GET -Uri $statsUrl -Headers $headers

                # Extract vulnerability counts by severity
                $CriticalCount = 0
                $HighCount = 0
                $MediumCount = 0
                $CriticalNP = 0
                $HighNP = 0
                $MediumNP = 0

                foreach ($severity in $stats.highSeverity, $stats.mediumSeverity, $stats.lowSeverity, $stats.infoSeverity) {
                    if ($severity) {
                        # Checkmarx uses High/Medium/Low/Info - map to our schema
                        # Assuming Critical maps to High in Checkmarx
                        switch ($severity.severityType) {
                            "High" { 
                                $CriticalCount = $severity.total
                                $CriticalNP = $severity.notExploitable
                            }
                            "Medium" { 
                                $HighCount = $severity.total
                                $HighNP = $severity.notExploitable
                            }
                            "Low" { 
                                $MediumCount = $severity.total
                                $MediumNP = $severity.notExploitable
                            }
                        }
                    }
                }

                # Get the most recent scan for this artifact to check if vulnerabilities changed
                $getPreviousScan = @"
SELECT ID, Critical, High, Medium, CriticalNP, HighNP, MediumNP, ScanRepeatCount, ScanDateTime 
FROM Scans 
WHERE ArtifactID = $artifactID 
  AND ScanTool = '$scanTool' 
  AND ScanType = '$scanType' 
ORDER BY ScanDateTime DESC 
LIMIT 1;
"@
                $previousScan = & $sqlitePath $dbPath $getPreviousScan
                
                if ($previousScan) {
                    # Parse the previous scan result
                    $prevFields = $previousScan -split '\|'
                    $prevID = $prevFields[0]
                    $prevCritical = [int]$prevFields[1]
                    $prevHigh = [int]$prevFields[2]
                    $prevMedium = [int]$prevFields[3]
                    $prevCriticalNP = [int]$prevFields[4]
                    $prevHighNP = [int]$prevFields[5]
                    $prevMediumNP = [int]$prevFields[6]
                    $prevRepeatCount = [int]$prevFields[7]
                    
                    # Check if all vulnerability counts match
                    if ($CriticalCount -eq $prevCritical -and $HighCount -eq $prevHigh -and $MediumCount -eq $prevMedium -and
                        $CriticalNP -eq $prevCriticalNP -and $HighNP -eq $prevHighNP -and $MediumNP -eq $prevMediumNP) {
                        # Same vulnerabilities - update date and increment repeat count
                        $updateScan = @"
UPDATE Scans 
SET ScanDateTime = '$DateForScanTable', 
    ScanRepeatCount = $($prevRepeatCount + 1) 
WHERE ID = $prevID;
"@
                        & $sqlitePath $dbPath $updateScan | Out-Null
                        Write-Verbose "Updated scan $prevID - repeat count now $($prevRepeatCount + 1)"
                    }
                    else {
                        # Different vulnerabilities - create new scan record
                        $insertScan = @"
PRAGMA foreign_keys = ON;
INSERT INTO Scans (ArtifactID, ScanTool, ScanType, ScanDateTime, ScanRepeatCount, Critical, High, Medium, CriticalNP, HighNP, MediumNP) 
VALUES ($artifactID, '$scanTool', '$scanType', '$DateForScanTable', 1, $CriticalCount, $HighCount, $MediumCount, $CriticalNP, $HighNP, $MediumNP);
"@
                        & $sqlitePath $dbPath $insertScan | Out-Null
                        Write-Verbose "Created new scan - vulnerabilities changed"
                    }
                }
                else {
                    # No previous scan - create new scan record
                    $insertScan = @"
PRAGMA foreign_keys = ON;
INSERT INTO Scans (ArtifactID, ScanTool, ScanType, ScanDateTime, ScanRepeatCount, Critical, High, Medium, CriticalNP, HighNP, MediumNP) 
VALUES ($artifactID, '$scanTool', '$scanType', '$DateForScanTable', 1, $CriticalCount, $HighCount, $MediumCount, $CriticalNP, $HighNP, $MediumNP);
"@
                    & $sqlitePath $dbPath $insertScan | Out-Null
                    Write-Verbose "Created first scan for Artifact $artifactID"
                }
            }

            # Track most recent scan for report
            if ($scanDate -gt $mostRecentDate) {
                $mostRecentDate = $scanDate
                $mostRecentScan = @{
                    Date = $DateForReport
                    Critical = $CriticalCount
                    High = $HighCount
                    Medium = $MediumCount
                }
            }
        }

        # Add the most recent scan to the report
        if ($mostRecentScan) {
            $details = [ordered]@{
                "Product Name" = $ProductName
                "Latest Scan Date (dd/MM/yyyy)" = $mostRecentScan.Date
                "Critical" = $mostRecentScan.Critical
                "High" = $mostRecentScan.High
                "Medium" = $mostRecentScan.Medium
            }
            $results += New-Object PSObject -Property $details
        }
    }

    # Export results to CSV
    $csvPath = "$PSScriptRoot\Checkmarx_Report.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation

    # Clean up old Excel file if it exists
    $oldExcelPath = Join-Path $GivePath_To_Generate_Report $filename
    if (Test-Path $oldExcelPath) {
        Remove-Item -Path $oldExcelPath -Force
    }

    # Convert CSV to Excel
    Import-Csv $csvPath | Export-Excel -Path (Join-Path $GivePath_To_Generate_Report $filename) -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    
    # Force garbage collection to release file handles
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Delete CSV file
    Remove-Item -Path $csvPath

    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Host "`nSuccessful!!!  $GivePath_To_Generate_Report$filename Duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"

}
catch {
    Write-Error $_.Exception.Message
    if ($_.Exception.Response) {
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $ErrResp = $streamReader.ReadToEnd()
        $streamReader.Close()
        Write-Error $ErrResp
    }
}
