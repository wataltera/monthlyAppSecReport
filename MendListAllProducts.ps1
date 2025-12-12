[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]
    [string]$userKey,
    [switch]$WD # Write details of a project in each product.  Write keys without this.
)
try {
    # Build the request body for getAllProducts
    $getAllProducts_body = @{
        "requestType" = "getAllProducts"
        "userKey"     = $userKey
        "orgToken"    = "2ba34226-c0b7-4c0a-aed7-d9e531b935f9"
    } | ConvertTo-Json

    # Call Mend.io API
    $productsUrl = "https://saas.mend.io/api/v1.4"
    $response = Invoke-RestMethod -Method Post `
        -Uri $productsUrl `
        -Body $getAllProducts_body `
        -ContentType "application/json"

    # $response | ConvertTo-Json -Depth 5 | Write-Output

    # Display product names and tokens
    foreach ($product in $response.products) {
        if ($WD) {
            Write-Output "Product: $($product.productName)"
        } else {
            Write-Output """$($product.productName)"" : ""$($product.productToken)"","
        }
        if ($WD) {
            try {
                # Call Mend API for projects inside this product
                $projectsUrl = "https://saas.mend.io/api/v1.4"
                    $projectsBody = @{
                    "requestType" = "getAllProjects"
                    "userKey" = $userKey
                    # "orgToken"    = "2ba34226-c0b7-4c0a-aed7-d9e531b935f9"
                    "productToken" = $product.productToken
                } | ConvertTo-Json

                # Write-Output "DEBUG: Request body:"
                # Write-Output $projectsBody

                $projectsResponse = Invoke-RestMethod -Method Post `
                    -Uri $projectsUrl `
                    -Body $projectsBody `
                    -ContentType "application/json"
                
                #Write-Output "DEBUG: Full response:"
                #$projectsResponse | ConvertTo-Json -Depth 5 | Write-Output
                
                if ($projectsResponse.projects) {
                    foreach ($project in $projectsResponse.projects) {
                        Write-Output "    Project: $($project.projectName)"
                    }
                } else {
                    Write-Output "No projects found in response"
                }
            }
            catch {
                Write-Output "DEBUG: Full error:"
                Write-Output $_.Exception | ConvertTo-Json -Depth 5
                Write-Output "DEBUG: Error details: $_"
                Write-Warning "Could not retrieve projects for $($product.productName): $_"
            }
        }

    }
}
catch {
    Write-Error "Error: $_"
}
