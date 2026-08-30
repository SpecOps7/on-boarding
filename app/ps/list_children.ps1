param([Parameter(Mandatory = $true)][string]$PathB64)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$path = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($PathB64))

$items = Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue | ForEach-Object {
    [PSCustomObject]@{
        name      = $_.Name
        is_dir    = $_.PSIsContainer
        size      = if ($_.PSIsContainer) { 0 } else { $_.Length }
        mtime_iso = $_.LastWriteTimeUtc.ToString('o')
        ext       = if ($_.PSIsContainer) { '' } else { $_.Extension.ToLower() }
    }
}

ConvertTo-Json -Compress -InputObject @($items)
