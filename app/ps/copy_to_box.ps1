# Copy one local file into a Box Drive folder (creates the destination folder).
param(
    [Parameter(Mandatory = $true)][string]$SrcB64,
    [Parameter(Mandatory = $true)][string]$DstB64
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$src = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($SrcB64))
$dst = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($DstB64))

try {
    if (-not (Test-Path -LiteralPath $src)) { throw "source not found: $src" }
    $dir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    Copy-Item -LiteralPath $src -Destination $dst -Force
    Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $true; dst = $dst; size = (Get-Item -LiteralPath $dst).Length })
}
catch {
    Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $false; dst = $dst; error = $_.Exception.Message })
}
