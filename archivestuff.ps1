$sourceDir = "C:\repo\Outlook\gpg2"
$archiveDir = "C:\repo\Outlook\archive\gpg2"

# Ensure the archive directory exists
if (!(Test-Path $archiveDir)) {
    New-Item -ItemType Directory -Force -Path $archiveDir
}

# Get all files and directories within the source directory created on or before 2024-01-01
Get-ChildItem -Path $sourceDir -Recurse | Where-Object {$_.CreationTime -lt (Get-Date -Year 2025 -Month 1 -Day 1)} | ForEach-Object {
    $targetPath = $_.FullName -replace [regex]::Escape($sourceDir), $archiveDir
    if ($_.PSIsContainer) {
       # It's a directory
       if (!(Test-Path $targetPath))
       {
        New-Item -ItemType Directory -Force -Path $targetPath
       }
       Move-Item -Path $_.FullName -Destination $targetPath -Force
    }
    else{
        # It's a file
        Move-Item -Path $_.FullName -Destination $targetPath -Force
    }
}

Write-Host "Files and folders created in 2024 or before have been moved to '$archiveDir'"