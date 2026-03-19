$sourceDir = "E:\DeThi"
Get-ChildItem -Path $sourceDir -Filter "*.docx" | ForEach-Object {
    if ($_.Name -match "^(\d+)-") {
        $num = [int]$matches[1]
        $start = [math]::Floor(($num - 1) / 100) * 100 + 1
        $end = $start + 99
        $targetFolder = Join-Path -Path $sourceDir -ChildPath "$start-$end"
        
        if (-not (Test-Path -Path $targetFolder)) {
            New-Item -ItemType Directory -Path $targetFolder | Out-Null
        }
        Move-Item -Path $_.FullName -Destination $targetFolder
    }
}
Write-Host "Mission failed successfully!"
