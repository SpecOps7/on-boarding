# Move Outlook messages into Inbox subfolders (created on demand) via COM.
# stdin lines: "<EntryID>|<base64 folder path relative to Inbox>"  — an empty path means the Inbox itself.
# If the target already holds a copy of the message (same Internet Message-ID — e.g. an earlier
# "copied instead of moved" during sync), the original is deleted instead of moved, so no duplicates.
# Emits one compact JSON line per message.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Emit($obj) { Write-Output (ConvertTo-Json -Compress -InputObject $obj) }

try { $ol = New-Object -ComObject Outlook.Application } catch { Emit @{ ok = $false; error = 'cannot attach to classic Outlook via COM' }; exit 0 }
$ns = $ol.GetNamespace('MAPI')
$inbox = $ns.GetDefaultFolder(6)
$cache = @{}
$MSGID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F'

function Get-OrCreateFolder($relPath) {
    if ($cache.ContainsKey($relPath)) { return $cache[$relPath] }
    $folder = $inbox
    foreach ($seg in ($relPath -split '\\')) {
        $seg = ($seg -replace '[\\/:*?"<>|]', '-').Trim()
        if (-not $seg) { continue }
        $child = $null
        foreach ($f in $folder.Folders) { if ($f.Name -eq $seg) { $child = $f; break } }
        if (-not $child) { $child = $folder.Folders.Add($seg) }
        $folder = $child
    }
    $cache[$relPath] = $folder
    return $folder
}

while ($null -ne ($line = [Console]::In.ReadLine())) {
    $line = $line.Trim()
    if (-not $line) { continue }
    $parts = $line.Split('|', 2)
    $entryId = $parts[0]
    $rel = if ($parts.Length -gt 1 -and $parts[1]) { [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($parts[1])) } else { '' }
    try {
        $item = $ns.GetItemFromID($entryId)
        $target = Get-OrCreateFolder $rel
        if ($item.Parent.EntryID -eq $target.EntryID) {
            Emit @{ ok = $true; entry_id = $entryId; folder = $target.FolderPath; skipped = $true }
            continue
        }
        $msgId = $null
        try { $msgId = $item.PropertyAccessor.GetProperty($MSGID) } catch {}
        if ($msgId) {
            $dupe = $null
            try { $dupe = $target.Items.Find("@SQL=""$MSGID"" = '" + ($msgId -replace "'", "''") + "'") } catch {}
            if ($dupe) {
                $item.Delete()
                Emit @{ ok = $true; entry_id = $entryId; new_entry_id = $dupe.EntryID; folder = $target.FolderPath; deduped = $true }
                continue
            }
        }
        $moved = $item.Move($target)
        # moving changes the EntryID — report the new one so the app can keep tracking the message
        Emit @{ ok = $true; entry_id = $entryId; new_entry_id = $moved.EntryID; folder = $target.FolderPath }
    }
    catch {
        Emit @{ ok = $false; entry_id = $entryId; error = $_.Exception.Message }
    }
}
