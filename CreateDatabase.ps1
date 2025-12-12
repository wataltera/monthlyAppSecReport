[CmdletBinding()]
param (
    [string]$DatabasePath = "C:\Users\W988276\AppData\Local\monthlyReportDatabase\monthlyReport.db",
    [string]$SqlitePath = "C:\sqlite\sqlite3.exe",
    [string]$SqlFilePath = "C:\Users\W988276\AppData\Local\monthlyReportDatabase\CreateArtifactsTable.sql"
)

# Verify paths exist
if (-not (Test-Path $SqlitePath)) {
    Write-Error "SQLite not found at: $SqlitePath"
    exit 1
}

if (-not (Test-Path $SqlFilePath)) {
    Write-Error "SQL file not found at: $SqlFilePath"
    exit 1
}

# Create database directory if it doesn't exist
$dbDir = Split-Path -Parent $DatabasePath
if (-not (Test-Path $dbDir)) {
    New-Item -ItemType Directory -Path $dbDir -Force | Out-Null
    Write-Host "Created directory: $dbDir"
}

# Create the table
Write-Host "Creating Artifacts table..." -ForegroundColor Green
$sqlContent = Get-Content $SqlFilePath -Raw
& $SqlitePath $DatabasePath $sqlContent

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ Table created successfully!" -ForegroundColor Green
} else {
    Write-Error "Failed to create table (exit code: $LASTEXITCODE)"
    exit 1
}

# Verify table was created
Write-Host "`nVerifying table structure..." -ForegroundColor Green
$verification = & $SqlitePath $DatabasePath "PRAGMA table_info(Artifacts);"
Write-Host $verification

# Show table info
Write-Host "`nTable summary:" -ForegroundColor Green
& $SqlitePath $DatabasePath ".schema Artifacts"

Write-Host "`n✓ Setup complete! Database ready at: $DatabasePath" -ForegroundColor Green
