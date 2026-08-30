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
    if (Test-Path -LiteralPath $dst) { throw "destination already exists: $dst" }
    Move-Item -LiteralPath $src -Destination $dst
    Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $true; src = $src; dst = $dst })
}
catch {
    Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $false; src = $src; dst = $dst; error = $_.Exception.Message })
}
