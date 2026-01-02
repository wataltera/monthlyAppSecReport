# This script retrieves Rapid7 InsightAppSec scan details and statistics for all applications organized by Business Unit tags.
# It requires the ImportExcel module for Excel operations.
# Ensure you have the ImportExcel module installed: Install-Module -Name ImportExcel -Scope CurrentUser

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ApiKey,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\powershell"
)

$DebugPreference = "SilentlyContinue"
$startTime = Get-Date

# Bypass SSL certificate validation if needed (use with caution)
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Determine OS
if ($PSVersionTable.Platform -eq "Unix" -and $PSVersionTable.OS -like "*Darwin*") {
    $mac = $true
} elseif ($PSVersionTable.Platform -eq "Win32NT" -or [string]::IsNullOrEmpty($PSVersionTable.Platform)) {
    $win = $true
} else {
    Write-Error "Can't determine OS from $($PSVersionTable.Platform), exiting"
    exit 1
}

# Define Rapid7 API URLs
# Note: Update the region if needed (us, eu, ca, au, ap)
$baseUrl = "https://ca.api.insight.rapid7.com/ias/v1"

# Output file paths
$today = Get-Date
$fileDate = $today.ToString("MM-dd-yyyy")
$timestamp = Get-Date -Format "MMM_dd_yy"

if ($mac) {
    $baseDir = "/Users/$env:USER/powershell/"
    $outputCsv = "$($baseDir)Rapid7_DAST.csv"
    $excelPath = "$($baseDir)Rapid7_DAST for $($fileDate).xlsx"
    $pwPath = "$($baseDir)APIKey.chm"
}
else {
    $baseDir = "C:\powershell\"
    $outputCsv = "$($baseDir)Rapid7_DAST.csv"
    $excelPath = "$($baseDir)Rapid7_DAST for $($fileDate).xlsx"
    $pwPath = "$($baseDir)APIKey.chm"
}
$outputCsv = Join-Path $OutputPath "Rapid7_DAST.csv"
$excelPath = Join-Path $OutputPath "Rapid7_DAST_for_$fileDate.xlsx"
$filename = "Rapid7_DAST_$timestamp.xlsx"

Write-Host "Output will be saved to: $(Join-Path $OutputPath $filename)" -ForegroundColor Cyan

# Clean up old files if they exist
if (Test-Path $outputCsv) {
    Remove-Item $outputCsv
}

# Set up authentication headers
$headers = @{
    "X-Api-Key" = $ApiKey
    "Accept" = "application/json"
}

try {
    $results = @()

    Write-Host "`n=== Step 1: Retrieving all tags from Rapid7 ===" -ForegroundColor Cyan
    
    # Get all tags from Rapid7 InsightAppSec
    $tagsUrl = "$baseUrl/tags"
    Write-Host "Calling: $tagsUrl" -ForegroundColor Gray
    
    try {
        $tagsResponse = Invoke-RestMethod -Uri $tagsUrl -Method Get -Headers $headers
        $allTags = $tagsResponse.data
        
        Write-Host "Retrieved $($allTags.Count) total tags" -ForegroundColor Green
        
        # Filter for tags that start with "BU:"
        $buTags = $allTags | Where-Object { $_.name -like "BU:*" }
        
        Write-Host "Found $($buTags.Count) Business Unit tags:" -ForegroundColor Green
        foreach ($tag in $buTags) {
            Write-Host "  - $($tag.name) (ID: $($tag.id))" -ForegroundColor Gray
        }
        
        if ($buTags.Count -eq 0) {
            Write-Host "No Business Unit tags found (tags starting with 'BU:')" -ForegroundColor Yellow
            Write-Host "Available tags:" -ForegroundColor Yellow
            $allTags | Select-Object -First 10 | ForEach-Object {
                Write-Host "  - $($_.name)" -ForegroundColor Gray
            }
            exit 0
        }
    }
    catch {
        Write-Error "Failed to retrieve tags: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            Write-Error "HTTP Status Code: $statusCode"
        }
        throw
    }
    
    Write-Host "`n=== Step 2: Retrieving all applications ===" -ForegroundColor Cyan
    
    # Get all applications from Rapid7
    $appsUrl = "$baseUrl/apps"
    Write-Host "Calling: $appsUrl" -ForegroundColor Gray
    
    try {
        $appsResponse = Invoke-RestMethod -Uri $appsUrl -Method Get -Headers $headers
        $allApplications = $appsResponse.data
        
        Write-Host "Retrieved $($allApplications.Count) total applications" -ForegroundColor Green
        
        if ($allApplications.Count -eq 0) {
            Write-Host "No applications found in Rapid7" -ForegroundColor Yellow
            exit 0
        }
    }
    catch {
        Write-Error "Failed to retrieve applications: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            Write-Error "HTTP Status Code: $statusCode"
        }
        throw
    }
    
    # Extract BU names from the tags for matching against app names
    $buNames = @()
    foreach ($tag in $buTags) {
        $buName = $tag.name -replace "^BU:", ""
        $buNames += $buName
    }
    Write-Host "Business Unit names to match: $($buNames -join ', ')" -ForegroundColor Cyan
    
    Write-Host "`n=== Step 3: Retrieving all scans ===" -ForegroundColor Cyan
    
    # Get all scans with pagination support
    $allScans = @()
    $pageIndex = 0
    $pageSize = 500  # Request maximum per page
    $hasMorePages = $true
    
    while ($hasMorePages) {
        $scansUrl = "$baseUrl/scans?size=$pageSize&index=$pageIndex"
        Write-Host "Calling: $scansUrl (page $($pageIndex + 1))" -ForegroundColor Gray
        
        try {
            $scansResponse = Invoke-RestMethod -Uri $scansUrl -Method Get -Headers $headers
            
            # Debug: Print pagination metadata
            if ($VerbosePreference -ne 'SilentlyContinue') {
                Write-Host "  Response metadata:" -ForegroundColor Yellow
                $scansResponse.metadata | ConvertTo-Json | Write-Host
            }
            
            $allScans += $scansResponse.data
            Write-Host "  Retrieved $($scansResponse.data.Count) scans (total so far: $($allScans.Count))" -ForegroundColor Gray
            
            # Check if there are more pages
            if ($scansResponse.metadata -and $scansResponse.metadata.total_pages) {
                $hasMorePages = $pageIndex -lt ($scansResponse.metadata.total_pages - 1)
            } elseif ($scansResponse.data.Count -lt $pageSize) {
                # No more data to fetch
                $hasMorePages = $false
            } else {
                # Assume more pages if we got a full page
                $hasMorePages = $true
            }
            
            $pageIndex++
            
            # Safety limit to prevent infinite loops
            if ($pageIndex -gt 100) {
                Write-Host "  Reached safety limit of 100 pages" -ForegroundColor Yellow
                $hasMorePages = $false
            }
        }
        catch {
            Write-Error "Failed to retrieve scans (page $pageIndex): $($_.Exception.Message)"
            if ($_.Exception.Response) {
                $statusCode = $_.Exception.Response.StatusCode.value__
                Write-Error "HTTP Status Code: $statusCode"
            }
            throw
        }
    }
    
    Write-Host "Retrieved $($allScans.Count) total scans" -ForegroundColor Green
    
    # Filter for completed scans only
    $completedScans = $allScans | Where-Object { $_.status -eq "COMPLETE" }
    Write-Host "Completed scans: $($completedScans.Count)" -ForegroundColor Green
    
    # Debug: Print JSON for completed scans (only if -Verbose)
    if ($VerbosePreference -ne 'SilentlyContinue') {
        Write-Host "`n--- Completed Scans JSON (first 3) ---" -ForegroundColor Yellow
        $completedScans | Select-Object -First 300 | ForEach-Object {
            $_ | ConvertTo-Json -Depth 5 | Write-Host
            Write-Host "---" -ForegroundColor Yellow
        }
    }
    
    if ($completedScans.Count -eq 0) {
        Write-Host "No completed scans found" -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "`n=== Step 4: Processing each application ===" -ForegroundColor Cyan
    
    foreach ($app in $allApplications) {
        Write-Host "`nProcessing: $($app.name) ID: $($app.id)" -ForegroundColor Gray
        
        # Determine Business Unit by checking if app name starts with any BU name
        $businessUnit = "Unknown"
        
        foreach ($buName in $buNames) {
            if ($app.name -like "$buName*") {
                $businessUnit = $buName
                Write-Host "  Matched BU: $businessUnit" -ForegroundColor Green
                break
            }
        }
        
        # Find all completed scans that belong to this application (using scan.app.id)
        $matchingScans = $completedScans | Where-Object { $_.app.id -eq $app.id }
        
        Write-Host "  Found $($matchingScans.Count) completed scan(s) for this app" -ForegroundColor Green
        
        if ($matchingScans.Count -eq 0) {
            Write-Host "  No scans found for " -ForegroundColor Yellow
            
            # Add a row to CSV indicating no scans found
            $noScanData = [ordered]@{
                "Business Unit" = $businessUnit
                "Application Name" = $app.name
                "Scan Date" = "No scans"
                "Critical" = 0
                "High" = 0
                "Medium" = 0
                "Low" = 0
            }
            $results += New-Object PSObject -Property $noScanData
            continue
        }
        
        # Add each matching scan to results
        foreach ($scan in $matchingScans) {
            Write-Host "    Processing scan ID: $($scan.id)" -ForegroundColor Cyan
            
            $scanDate = "Unknown"
            if ($scan.completion_time) {
                $scanDate = Get-Date $scan.completion_time -Format "yyyy-MM-dd HH:mm:ss"
            }
            
            # Get vulnerability counts
            $critical = 0
            $high = 0
            $medium = 0
            $low = 0
            
            # Try to get vulnerability counts from scan object or via API
            if ($scan.vulnerability_score) {
                $critical = if ($scan.vulnerability_score.critical) { $scan.vulnerability_score.critical } else { 0 }
                $high = if ($scan.vulnerability_score.high) { $scan.vulnerability_score.high } else { 0 }
                $medium = if ($scan.vulnerability_score.medium) { $scan.vulnerability_score.medium } else { 0 }
                $low = if ($scan.vulnerability_score.low) { $scan.vulnerability_score.low } else { 0 }
            } else {
                # Get vulnerabilities for this scan using search API
                $searchUrl = "$baseUrl/search"
                $searchBody = @{
                    "type" = "VULNERABILITY"
                    "query" = "vulnerability.scans.id='$($scan.id)'"
                } | ConvertTo-Json
                # The lines above got a 415 return code, using brute force method
                $searchBody = "{""type"":""VULNERABILITY"",""query"":""vulnerability.scans.id='$($scan.id)'""}"
                if ($VerbosePreference -ne 'SilentlyContinue') {
                    Write-Host $searchBody
                }

                try {
                    $vulnResponse = Invoke-RestMethod -Uri $searchUrl -Method Post -Headers $headers -ContentType "application/json"-Body $searchBody
                    
                    # Debug: Print the vulnerability response JSON (only if -Verbose)
                    if ($VerbosePreference -ne 'SilentlyContinue') {
                        Write-Host "    --- Vulnerability Response JSON ---" -ForegroundColor Yellow
                        $vulnResponse | ConvertTo-Json -Depth 5 | Write-Host
                        Write-Host "    --- End JSON ---" -ForegroundColor Yellow
                    }
                    
                    # Count vulnerabilities by severity
                    foreach ($vuln in $vulnResponse.data) {
                        switch ($vuln.severity) {
                            "CRITICAL" { $critical++ }
                            "HIGH" { $high++ }
                            "MEDIUM" { $medium++ }
                            "LOW" { $low++ }
                        }
                    }
                }
                catch {
                    Write-Host "    Could not retrieve vulnerabilities for scan $($scan.id): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            
            Write-Host "    Scan $($scan.id): $scanDate - C:$critical H:$high M:$medium L:$low" -ForegroundColor Cyan
            
            # Add data structure for CSV output
            $scanData = [ordered]@{
                "Business Unit" = $businessUnit
                "Application Name" = $app.name
                "Scan Date" = $scanDate
                "Critical" = $critical
                "High" = $high
                "Medium" = $medium
                "Low" = $low
            }
            
            $results += New-Object PSObject -Property $scanData
        }
    }
    
    Write-Host "`n=== Step 5: Exporting results ===" -ForegroundColor Cyan
    
    if ($results.Count -gt 0) {
        # Export to CSV
        $results | Export-Csv -Path $outputCsv -NoTypeInformation
        Write-Host "CSV exported to: $outputCsv" -ForegroundColor Green
        
        # Export to Excel with formatting
        $results | Export-Excel -Path (Join-Path $OutputPath $filename) `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow `
            -WorksheetName "Rapid7 DAST"
        
        Write-Host "Excel exported to: $(Join-Path $OutputPath $filename)" -ForegroundColor Green
    }
    else {
        Write-Host "No results to export" -ForegroundColor Yellow
    }
    
    # Clean up CSV file
    if (Test-Path $outputCsv) {
        Remove-Item $outputCsv
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-Host "`nSuccessful!!! Duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor Green
}
catch {
    Write-Error "Error occurred: $($_.Exception.Message)"
    Write-Error $_.Exception
    exit 1
}
