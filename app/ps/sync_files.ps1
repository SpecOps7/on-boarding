# Reads "srcB64|dstB64" pairs (base64 UTF-8 Windows file paths) from stdin, one per line.
# Copies each file (hydrates Box placeholders), emits one compact JSON line per file.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

while ($null -ne ($line = [Console]::In.ReadLine())) {
    $line = $line.Trim()
    if (-not $line) { continue }
    $parts = $line.Split('|')
    $src = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($parts[0]))
    $dst = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($parts[1]))
    try {
        $dir = Split-Path -Parent $dst
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $s = Get-Item -LiteralPath $src
        $skip = $false
        if (Test-Path -LiteralPath $dst) {
            $d = Get-Item -LiteralPath $dst
            if ($d.Length -eq $s.Length -and $d.LastWriteTimeUtc -eq $s.LastWriteTimeUtc) { $skip = $true }
        }
        if (-not $skip) {
            Copy-Item -LiteralPath $src -Destination $dst -Force
            (Get-Item -LiteralPath $dst).LastWriteTimeUtc = $s.LastWriteTimeUtc
        }
        Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $true; src = $src; skipped = $skip })
    }
    catch {
        Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $false; src = $src; error = $_.Exception.Message })
    }
}
