[CmdletBinding()]
param (
    # Verbose is a built-in PowerShell parameter which controls Write-Verbose output
    # *> redirects everything includeing write host.  > just redirects Write-Output
    # -Verbose 4>&1 > combined.txt includes Write-Verbose and Write-Output
    [Parameter(Mandatory = $True)]
    [string]$GivePath_To_Generate_Report,

    [Parameter(Mandatory = $True)]
    [string]$userKey
)
$startTime = Get-Date

$scanTool = "Mend";
$scanType = "SCA";
$orgToken = "2ba34226-c0b7-4c0a-aed7-d9e531b935f9"

<#
Mend dynamically discovers all products and projects from the organization.
This program collects vulnerability statistics for all projects across all products.
#>
# Build filename timestamp in MMM dd YY format
$timestamp = Get-Date -Format "MMM_dd_yy"
# Build filename: AllProducts_Mend_SCA_MMM_dd_yy.xlsx
$filename = "AllProducts_${scanTool}_${scanType}_${timestamp}.xlsx"
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

    # Get all products from Mend using getAllProducts API
    Write-Host "Discovering all products from Mend..." -ForegroundColor Cyan
    
    $getAllProducts_body = @{
        "requestType" = "getAllProducts"
        "userKey"     = $userKey
        "orgToken"    = $orgToken
    } | ConvertTo-Json

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $getAllProducts_response = Invoke-WebRequest -Method POST -ContentType 'application/json' -Body $getAllProducts_body -Uri "https://saas.mend.io/api/v1.4"
    $getAllProducts_content = $getAllProducts_response.Content | ConvertFrom-Json
    $allProducts = $getAllProducts_content.products

    Write-Host "Found $($allProducts.Count) products" -ForegroundColor Green

    # Iterate through all products
    foreach ($product in $allProducts) {
        $ProductName = $product.productName
        $ProductToken = $product.productToken

        Write-Host "Collecting Data for $ProductName ..." 

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

            Write-Output "`nWorking on $ProductName - $projectName"
            Write-Host "`nWorking on $ProductName - $projectName"
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
                Write-Output "No vulnerabilities found in $ProductName $projectName"
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
    Write-Host "`nSuccessful!!!  $(Join-Path $GivePath_To_Generate_Report $filename) Duration: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s"
} 
catch {
    Write-Error $_.Exception
    $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
    $ErrResp = $streamReader.ReadToEnd() | ConvertFrom-Json
    $streamReader.Close()
    $ErrResp
}
