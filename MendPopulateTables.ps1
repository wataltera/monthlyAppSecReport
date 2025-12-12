[CmdletBinding()]
param (
    # Verbose is a built-in PowerShell parameter which controls Write-Verbose output
    # *> redirects everything includeing write host.  > just redirects Write-Output
    # -Verbose 4>&1 > combined.txt includes Write-Verbose and Write-Output
    [Parameter(Mandatory = $True)]
    [string]$GivePath_To_Generate_Report,

    [Parameter(Mandatory = $True)]
    [string]$userKey,

    [string]$whichBU   # choose which group of products to load
)
$startTime = Get-Date

$scanTool = "Mend";
$scanType = "SCA";

<#
Mend stores only one scan for each combination of product and project.
This program populates the Artifacts table if there is no Artifact record
then adds a new scan row in the Scans table if there is no scan
on that date.
#>
# Build filename timestamp in MMM dd YY format
$timestamp = Get-Date -Format "MMM_dd_yy"
# Build filename: GroupName_Mend_SCA_MMM_dd_yy.xlsx
$filename = "${whichBU}_${scanTool}_${scanType}_${timestamp}.xlsx"
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

    # Build paths for config files
    $jsonTemplatePath = Join-Path $PSScriptRoot "ProductGroups.json"
    $jsonLocalPath = Join-Path $PSScriptRoot "ProductGroups.local.json"

    # Try to use local config first (with real tokens), fall back to template
    # ProductGroups.local.json should be gitignored and contain real tokens
    # ProductGroups.json is the template in version control
    if (Test-Path $jsonLocalPath) {
        $configPath = $jsonLocalPath
        Write-Host "Using local config: $jsonLocalPath" -ForegroundColor Gray
    }
    elseif (Test-Path $jsonTemplatePath) {
        $configPath = $jsonTemplatePath
        Write-Host "Using template config: $jsonTemplatePath (consider creating ProductGroups.local.json with real tokens)" 
    }
    else {
        throw "Config files not found. Create either:$([Environment]::NewLine)  - $jsonTemplatePath (template)$([Environment]::NewLine)  - $jsonLocalPath (local with real tokens, gitignored)"
    }

    # Read JSON file containing all product groups
    $groups = Get-Content $configPath | ConvertFrom-Json

    # Select the group requested on the command line (case-sensitive match)
    $matchedProperty = $groups.PSObject.Properties | Where-Object { $_.Name -ceq $whichBU }
    
    if (-not $matchedProperty) {
        throw "Group '$whichBU' not found in $configPath (case-sensitive match). Available groups: $($groups.PSObject.Properties.Name -join ', ')"
    }

    # Get Mend products for this BU
    $mendProperty = $matchedProperty.Value.PSObject.Properties | Where-Object { $_.Name -eq "Mend" }
    
    if (-not $mendProperty) {
        throw "No Mend products defined for BU '$whichBU' in $configPath"
    }

    # Convert it to a true hash table to be compatible with the original program
    $ProductToken = @{}
    $mendProperty.Value.PSObject.Properties | ForEach-Object {
        $ProductToken[$_.Name] = $_.Value
    }

    # Example: show what was loaded
    foreach ($entry in $ProductToken.GetEnumerator()) {
        Write-Verbose "Product: $($entry.Key)  Token: $($entry.Value)"
    }

    $ProductToken.GetEnumerator() | ForEach-Object {
        $ProductName =  $_.Key
        $ProductToken = $_.value

        Write-Host "Collecting Data of $whichBU $ProductName ..." 

        $getAllProjects_body= @{
            "requestType"  = "getAllProjects"
            "userKey"      = $userKey
            "productToken" = $ProductToken
        } | ConvertTo-Json

        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $getAllProjects_response = Invoke-WebRequest -Method POST -ContentType 'application/json' -Body $getAllProjects_body -Uri "https://saas.mend.io/api/v1.4"
        $getAllProjects_response_content = $getAllProjects_response.Content | ConvertFrom-Json
        $projects = $getAllProjects_response_content.projects

        #---------------------------------------------------------------------------------------

        foreach($project in $projects)
        {
            $projectName = $project.projectName
            $projectToken = $project.projectToken
            $MediumCount = 0
            $HighCount = 0
            $CriticalCount = 0

            Write-Output "`nWorking on $whichBU $projectName"
            Write-Host "`nWorking on $whichBU $projectName"
            #Get "Last Scan Date"
            $ProjectScanLastDate_Body = @{
                "requestType"  = "getProjectVitals"
                "userKey"      = "$userKey"
                "projectToken" = "$projectToken"
            } | ConvertTo-Json

            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $ProjectScanLastDate_response = Invoke-WebRequest -Method POST -ContentType 'application/json' -Body $ProjectScanLastDate_Body -Uri "https://saas.mend.io/api/v1.4"
            $Project_Vital_content = $ProjectScanLastDate_response.Content | ConvertFrom-Json
            $Date = Get-Date($Project_Vital_content.projectVitals.lastUpdatedDate) -Format "MMM dd, yyyy hh:mm:ss tt"
            # Generate sortable date for database (uppercase MM for month, not lowercase mm for minutes)
            $DateForScanTable = Get-Date($Project_Vital_content.projectVitals.lastUpdatedDate) -Format "yyyy-MM-dd HH:mm:ss"
            # Conditionally dump the project metadata
            if ($VerbosePreference -ne 'SilentlyContinue') {
                Write-Verbose "Dates for $projectName $Date $DateForScanTable"
                $Project_Vital_content | ConvertTo-Json -Depth 10 | Write-Verbose
            }
            #GetProject Alerts (Total Count of vulnerability)
            $body= @{
                "requestType"  = "getProjectAlerts"
                "userKey"      = $userKey
                "projectToken" = $projectToken
                # "productToken" = $ProductToken
            } | ConvertTo-Json

            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $response = Invoke-WebRequest -Method POST -ContentType 'application/json' -Body $body -Uri "https://saas.mend.io/api/v1.4"
            $content = $response.Content | ConvertFrom-Json
            # alerts[] may be empty
            $vulnerabilities_Score = $content.alerts.vulnerability.cvss3_severity
    
            # Dump vulnerabilities found in the scan if desired
            if ($VerbosePreference -ne 'SilentlyContinue') {
                $content | ConvertTo-Json -Depth 10 | Write-Verbose
            }

            if ($vulnerabilities_Score) {
                Write-Output "Vulnerabilities found: $($vulnerabilities_Score.Count)"
                foreach ($score in $vulnerabilities_Score) {
                    if ($score -eq "medium") { $MediumCount++ }
                    if ($score -eq "high") { $HighCount++ }
                    if ($score -eq "critical") { $CriticalCount++ }
                }
            }
            else {
                # No vulnerabilities for this product / project
                Write-Output "No vulnerabilities found in $ProductName $projectName ArtifactID $artifactID"
                $CriticalCount = -1
                $HighCount = -1
                $MediumCount = -1
            }

            $details = [ordered]@{            
                "Product Name"                   = $ProductName  
                "Project Name"                   = $projectName
                "Latest Scan Date"              = $Date 
                "Critical"                       = $CriticalCount
                "High"                           = $HighCount 
                "Medium"                         = $MediumCount
            }  
            # Collect the scan data into a .csv record for later writing to Excel
            $results += New-Object PSObject -Property $details 
            $results | export-csv -Path "$PSScriptRoot\WhiteSource_Report.csv" -NoTypeInformation
            
            # Populate the Artifacts and Scans tables
            $dbPath = "C:\Users\W988276\AppData\Local\monthlyReportDatabase\monthlyReport.db"
            $sqlitePath = "C:\sqlite\sqlite3.exe"
            
            # Enable foreign key constraints for this connection
            & $sqlitePath $dbPath "PRAGMA foreign_keys = ON;"
            
            # Get or create Artifact ID
            $query = @"
SELECT ID FROM Artifacts 
WHERE MendProduct = '$($ProductName.Replace("'", "''"))' 
  AND MendProject = '$($projectName.Replace("'", "''"))' 
  AND BusinessUnit = '$($whichBU.Replace("'", "''"))';
"@
            $artifactID = & $sqlitePath $dbPath $query
            
            if (-not $artifactID) {
                # Create new Artifact record
                $insertArtifact = @"
PRAGMA foreign_keys = ON;
INSERT INTO Artifacts (BusinessUnit, MendProduct, MendProject, SCAScans, RecentSCA) 
VALUES ('$($whichBU.Replace("'", "''"))', '$($ProductName.Replace("'", "''"))', '$($projectName.Replace("'", "''"))', 0, NULL);
SELECT last_insert_rowid();
"@
                $artifactID = & $sqlitePath $dbPath $insertArtifact
                Write-Verbose "Created new Artifact ID: $artifactID"
            }
            else {
                Write-Verbose "Found existing Artifact ID: $artifactID"
            }
            
            # Check if scan already exists for this artifact, tool, and type on this date
            $scanTool = "Mend"
            $scanType = "SCA"
            
            Write-Debug "DEBUG: Checking for existing scan - Artifact:$artifactID Tool:$scanTool Type:$scanType Date:$DateForScanTable" 
            
            $checkScan = @"
SELECT ID FROM Scans 
WHERE ArtifactID = $artifactID 
  AND ScanTool = '$scanTool' 
  AND ScanType = '$scanType' 
  AND ScanDateTime = '$DateForScanTable';
"@
            $existingScanID = & $sqlitePath $dbPath $checkScan
            
            if ($existingScanID) {
                Write-Debug "DEBUG: Scan $existingScanID already exists - skipping" 
                continue
            }
            
            Write-Debug "DEBUG: No existing scan found, checking for previous scan to compare vulnerabilities" 
            
            # Get the most recent scan for this artifact/tool/type to check if vulnerabilities changeds
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
            
            # Note: Mend only provides Critical, High, Medium (no NP variants in current data)
            # Set NP values to 0 for now
            $CriticalNP = 0
            $HighNP = 0
            $MediumNP = 0
            
            if ($previousScan) {
                # Parse the previous scan result (format: "ID|Critical|High|Medium|CriticalNP|HighNP|MediumNP|ScanRepeatCount")
                $prevFields = $previousScan -split '\|'
                $prevID = $prevFields[0]
                $prevCritical = [int]$prevFields[1]
                $prevHigh = [int]$prevFields[2]
                $prevMedium = [int]$prevFields[3]
                $prevCriticalNP = [int]$prevFields[4]
                $prevHighNP = [int]$prevFields[5]
                $prevMediumNP = [int]$prevFields[6]
                $prevRepeatCount = [int]$prevFields[7]
                $prevScanDateTime = $prevFields[8]
                
                # Check if all vulnerability counts match
                if ($CriticalCount -eq $prevCritical -and $HighCount -eq $prevHigh -and $MediumCount -eq $prevMedium -and
                    $CriticalNP -eq $prevCriticalNP -and $HighNP -eq $prevHighNP -and $MediumNP -eq $prevMediumNP) {
                    # Same vulnerabilities - update date and increment repeat count, no effect on foreign keys
                    Write-Debug "DEBUG: Vulnerabilities match previous scan $prevID (prev date: $prevScanDateTime, new date: $DateForScanTable)" 
                    Write-Debug "DEBUG: About to UPDATE scan $prevID to new date $DateForScanTable" 
                    
                    # Check if the new date would conflict with any existing scan
                    $checkConflict = @"
SELECT ID FROM Scans 
WHERE ArtifactID = $artifactID
  AND ScanTool = '$scanTool' 
  AND ScanType = '$scanType' 
  AND ScanDateTime = '$DateForScanTable'
  AND ID != $prevID;
"@
                    $conflictID = & $sqlitePath $dbPath $checkConflict
                    
                    if ($conflictID) {
                        Write-Host "ERROR: Cannot update scan $prevID to date $DateForScanTable - conflicts with existing scan $conflictID" 
                        Write-Host "Skipping this scan update" 
                        continue
                    }
                    
                    $updateScan = @"
PRAGMA foreign_keys = ON;
UPDATE Scans 
SET ScanDateTime = '$DateForScanTable', 
    ScanRepeatCount = $($prevRepeatCount + 1) 
WHERE ID = $prevID;
"@
                    & $sqlitePath $dbPath $updateScan | Out-Null
                    Write-Debug "DEBUG: Successfully updated scan $prevID - repeat count now $($prevRepeatCount + 1)" 
                }
                else {
                    # Different vulnerabilities - create new scan record
                    Write-Debug "DEBUG: Vulnerabilities changed - inserting new scan (C:$CriticalCount H:$HighCount M:$MediumCount vs prev C:$prevCritical H:$prevHigh M:$prevMedium)" 
                    $insertScan = @"
PRAGMA foreign_keys = ON;
INSERT INTO Scans (ArtifactID, ScanTool, ScanType, ScanDateTime, ScanRepeatCount, Critical, High, Medium, CriticalNP, HighNP, MediumNP) 
VALUES ($artifactID, '$scanTool', '$scanType', '$DateForScanTable', 1, $CriticalCount, $HighCount, $MediumCount, $CriticalNP, $HighNP, $MediumNP);
"@
                    $insertResult = & $sqlitePath $dbPath $insertScan 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Host "ERROR inserting scan: $insertResult" 
                        throw "Failed to insert scan"
                    }
                    Write-Debug "DEBUG: Successfully created new scan" 
                }
            }
            else {
                # No previous scan - create new scan record
                Write-Debug "DEBUG: No previous scan found - inserting scan" 
                $insertScan = @"
PRAGMA foreign_keys = ON;
INSERT INTO Scans (ArtifactID, ScanTool, ScanType, ScanDateTime, ScanRepeatCount, Critical, High, Medium, CriticalNP, HighNP, MediumNP) 
VALUES ($artifactID, '$scanTool', '$scanType', '$DateForScanTable', 1, $CriticalCount, $HighCount, $MediumCount, $CriticalNP, $HighNP, $MediumNP);
"@
                $insertResult = & $sqlitePath $dbPath $insertScan 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "ERROR inserting scan: $insertResult" 
                    throw "Failed to insert scan $artifactID $scanTool $scanType $DateForScanTable"
                }
                Write-Debug "DEBUG: Successfully created first scan for Artifact $artifactID" 
            }

        }
    }

    # -----------------------------------------------------------
    # -----------------------------------------------------------
    $File_Total_Drop = Get-ChildItem $GivePath_To_Generate_Report | Select-Object FullName
    foreach($item in $File_Total_Drop){
        $File_Fullname = $item.FullName
        if ($File_Fullname -eq "$GivePath_To_Generate_Report\WhiteSource_Report.xlsx") {
            Remove-Item -Path $File_Fullname -Force
        }
    }

    #  # **********************START*************************************Convert CSV to EXCEL********************************************
    

    Import-Csv "$PSScriptRoot\WhiteSource_Report.csv" | Export-Excel -Path (Join-Path $GivePath_To_Generate_Report $filename) -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    
    # Force garbage collection to release file handles  The Excel export sometimes leaves the file open
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    #  # **********************END*************************************Convert CSV to EXCEL********************************************
      
    # # **********************START*************************************Delete CSV FILE********************************************
    $csv_path= (Get-ChildItem "$PSScriptRoot\" -Filter "WhiteSource_Report.csv").FullName
    Remove-Item -Path "$csv_path"
   # # **********************END*************************************Delete CSV FILE********************************************

    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Host "`nSuccessful!!!  $GivePath_To_Generate_Report$filename Duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
} 
catch {
    Write-Error $_.Exception
    $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
    $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
    $streamReader.Close()
    $ErrResp
}
