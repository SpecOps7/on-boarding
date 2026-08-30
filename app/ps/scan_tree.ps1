# Reads base64(UTF-8) Windows folder paths from stdin, one per line.
# For each, emits ONE compact JSON line: {path, files:[{n,s,m}]} or {path, error}.
# Listing metadata only — never opens file contents (no Box hydration).
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

while ($null -ne ($line = [Console]::In.ReadLine())) {
    $line = $line.Trim()
    if (-not $line) { continue }
    $path = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line))
    try {
        $files = @(Get-ChildItem -LiteralPath $path -Recurse -Depth 6 -File -Force -ErrorAction SilentlyContinue |
            ForEach-Object {
                @{ n = $_.FullName.Substring($path.Length).TrimStart('\'); s = $_.Length; m = $_.LastWriteTimeUtc.ToString('o') }
            })
        Write-Output (ConvertTo-Json -Compress -Depth 4 -InputObject @{ path = $path; files = $files })
    }
    catch {
        Write-Output (ConvertTo-Json -Compress -InputObject @{ path = $path; error = $_.Exception.Message })
    }
}
