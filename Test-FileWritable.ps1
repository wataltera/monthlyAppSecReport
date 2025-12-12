function Test-FileWritable {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        # Try to open the file with write access
        $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'Write', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        return $true
    }
    catch {
        Write-Error "Cannot write to file '$FilePath'. It may be open in another program (like Excel) or you may lack permissions. Error: $($_.Exception.Message)"
        return $false
    }
}

# Example usage:
# if (-not (Test-FileWritable $filename)) {
#     exit 1
# }
