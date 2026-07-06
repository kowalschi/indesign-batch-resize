$path = "D:\git\history_redimensionari"
$commands = @()

Get-ChildItem -Path $path -File | ForEach-Object {
    $date = $_.LastWriteTime
    $year = $date.Year
    $yearFolder = Join-Path $path $year

    if (-not (Test-Path $yearFolder)) {
        New-Item -ItemType Directory -Path $yearFolder | Out-Null
    }

    $destination = Join-Path $yearFolder $_.Name
    Move-Item -Path $_.FullName -Destination $destination

    $dateStr = $date.ToString("yyyy-MM-ddTHH:mm:ss")
    $relativePath = "$year/$($_.Name)"

    $commands += "git add `"$relativePath`""
    $commands += "git commit --date=`"$dateStr`" -m `"Add $($_.Name)`""
    $commands += ""
}

$commands | Out-File -FilePath "$path\git_commands.txt" -Encoding UTF8

Write-Host "Done! Files organized and git commands saved to git_commands.txt"