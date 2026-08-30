param(
    [Parameter(Mandatory = $true)][string]$SrcB64,
    [Parameter(Mandatory = $true)][string]$DstB64
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$src = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($SrcB64))
$dst = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($DstB64))

$excludedExts = @('.indd', '.psd', '.ai', '.mp4', '.mov', '.zip', '.exe', '.dll')
$maxBytes = 50MB

New-Item -ItemType Directory -Force -Path $dst | Out-Null

$files = @(Get-ChildItem -LiteralPath $src -Recurse -File -Force -ErrorAction SilentlyContinue)
$total = $files.Count
Write-Output (ConvertTo-Json -Compress -InputObject @{ event = 'start'; total = $total })

$copied = 0; $skipped = 0; $excluded = 0; $errors = 0
foreach ($f in $files) {
    $rel = $f.FullName.Substring($src.Length).TrimStart('\')
    try {
        $ext = $f.Extension.ToLower()
        if ($excludedExts -contains $ext -or $f.Length -gt $maxBytes) {
            $excluded++
            Write-Output (ConvertTo-Json -Compress -InputObject @{
                event = 'file'; file = $rel; action = 'excluded'
                size = $f.Length; mtime = $f.LastWriteTimeUtc.ToString('o'); ext = $ext
            })
            continue
        }

        $destPath = Join-Path $dst $rel
        $destDir = Split-Path -Parent $destPath
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Force -Path $destDir | Out-Null
        }

        $needCopy = $true
        if (Test-Path -LiteralPath $destPath) {
            $d = Get-Item -LiteralPath $destPath
            if ($d.Length -eq $f.Length -and $d.LastWriteTimeUtc -eq $f.LastWriteTimeUtc) {
                $needCopy = $false
            }
        }

        if ($needCopy) {
            # Copy-Item hydrates the Box placeholder, then we pin mtime for incremental compares
            Copy-Item -LiteralPath $f.FullName -Destination $destPath -Force
            (Get-Item -LiteralPath $destPath).LastWriteTimeUtc = $f.LastWriteTimeUtc
            $copied++
            $action = 'copied'
        }
        else {
            $skipped++
            $action = 'skipped'
        }

        Write-Output (ConvertTo-Json -Compress -InputObject @{
            event = 'file'; file = $rel; action = $action
            size = $f.Length; mtime = $f.LastWriteTimeUtc.ToString('o'); ext = $ext
        })
    }
    catch {
        $errors++
        Write-Output (ConvertTo-Json -Compress -InputObject @{
            event = 'file'; file = $rel; action = 'error'; message = $_.Exception.Message
        })
    }
}

Write-Output (ConvertTo-Json -Compress -InputObject @{
    event = 'done'; total = $total; copied = $copied
    skipped = $skipped; excluded = $excluded; errors = $errors
})
