# Pull mail from classic Outlook via COM (no API / app registration).
# Emits one compact JSON line per message, then {"event":"done",...}.
param(
    [Parameter(Mandatory = $true)][string]$SinceB64,      # ISO-8601 watermark (local time)
    [Parameter(Mandatory = $true)][string]$AttachDirB64,  # Windows dir for saved attachments
    [string]$Folders = 'Inbox,Sent',
    [int]$MaxBody = 20000
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$since = [datetime]::Parse([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($SinceB64)))
$attachDir = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($AttachDirB64))
New-Item -ItemType Directory -Force -Path $attachDir | Out-Null

function Emit($obj) { Write-Output (ConvertTo-Json -Compress -Depth 5 -InputObject $obj) }

# --- attach to Outlook; auto-start it (minimized) if COM can't reach a running instance
$ol = $null
for ($attempt = 1; $attempt -le 3 -and -not $ol; $attempt++) {
    try { $ol = New-Object -ComObject Outlook.Application }
    catch {
        if ($attempt -eq 1) {
            Emit @{ event = 'info'; message = 'Outlook not reachable via COM - starting classic Outlook' }
            try { Start-Process 'outlook.exe' -WindowStyle Minimized } catch {}
        }
        Start-Sleep -Seconds 20
    }
}
if (-not $ol) { Emit @{ event = 'error'; message = 'could not attach to classic Outlook via COM (0x80080005). Is it installed and able to start?' }; exit 0 }

$ns = $ol.GetNamespace('MAPI')
if ($ns.Accounts.Count -eq 0) {
    Emit @{ event = 'error'; message = 'classic Outlook has no mail account - add the account (File > Add Account) and let it sync' }
    exit 0
}

$folderIds = @{ 'Inbox' = 6; 'Sent' = 5 }
$saveExts = @('.pdf', '.docx', '.doc', '.xlsx', '.xls', '.msg')
$count = 0; $maxReceived = $since
$filterDate = $since.ToString('MM/dd/yyyy HH:mm')
$seen = New-Object 'System.Collections.Generic.HashSet[string]'   # live sync can re-enumerate items

foreach ($name in ($Folders -split ',')) {
    $name = $name.Trim()
    if (-not $folderIds.ContainsKey($name)) { continue }
    try { $folder = $ns.GetDefaultFolder($folderIds[$name]) } catch { Emit @{ event = 'warn'; message = "no $name folder" }; continue }
    $field = if ($name -eq 'Sent') { '[SentOn]' } else { '[ReceivedTime]' }
    $items = $folder.Items
    $items.Sort($field, $false)
    $items = $items.Restrict("$field >= '$filterDate'")

    foreach ($m in $items) {
        try {
            if ($m.Class -ne 43) { continue }   # olMail only
            if (-not $seen.Add($m.EntryID)) { continue }
            $when = if ($name -eq 'Sent') { $m.SentOn } else { $m.ReceivedTime }
            if ($when -gt $maxReceived) { $maxReceived = $when }

            $smtp = ''
            try { $smtp = $m.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E') } catch {}
            if (-not $smtp) { try { $smtp = $m.SenderEmailAddress } catch {} }

            $body = ''
            try { $body = $m.Body } catch {}
            if ($body.Length -gt $MaxBody) { $body = $body.Substring(0, $MaxBody) }

            $atts = @()
            $idKey = $m.EntryID.Substring([Math]::Max(0, $m.EntryID.Length - 16))
            foreach ($a in $m.Attachments) {
                $ext = [System.IO.Path]::GetExtension($a.FileName).ToLower()
                $saved = $null
                if ($saveExts -contains $ext -and $a.Size -lt 30MB) {
                    try {
                        $dir = Join-Path $attachDir $idKey
                        New-Item -ItemType Directory -Force -Path $dir | Out-Null
                        $safe = ($a.FileName -replace '[\\/:*?"<>|]', '_')
                        $saved = Join-Path $dir $safe
                        if (-not (Test-Path -LiteralPath $saved)) { $a.SaveAsFile($saved) }
                    } catch { $saved = $null }
                }
                $atts += @{ name = $a.FileName; size = $a.Size; saved_path = $saved }
            }

            Emit @{
                event = 'mail'; entry_id = $m.EntryID; conversation_id = $m.ConversationID
                folder = $name; received = $when.ToString('o')
                sender_name = $m.SenderName; sender_email = $smtp
                to = $m.To; cc = $m.CC; subject = $m.Subject
                body = $body; attachments = $atts
            }
            $count++
        }
        catch { Emit @{ event = 'warn'; message = ('skipped item: ' + $_.Exception.Message) } }
    }
}

Emit @{ event = 'done'; count = $count; max_received = $maxReceived.ToString('o') }
