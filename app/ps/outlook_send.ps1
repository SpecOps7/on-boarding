# Send an email (or save it to Drafts) via classic Outlook COM.
# -SpecB64: base64 UTF-8 JSON {mode: "send"|"draft", to, cc, subject, body, reply_entry_id}
# If reply_entry_id is set, replies within that message's thread (keeps history + threading).
# Emits one compact JSON result line.
param([Parameter(Mandatory = $true)][string]$SpecB64)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$spec = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($SpecB64)) | ConvertFrom-Json

try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    if ($ns.Accounts.Count -eq 0) { throw 'classic Outlook has no mail account' }

    $item = $null
    if ($spec.reply_entry_id) {
        try {
            $orig = $ns.GetItemFromID($spec.reply_entry_id)
            $item = $orig.Reply()
            # reply body: our text on top, quoted history below
            $item.Body = $spec.body + "`r`n`r`n" + $item.Body
        } catch { $item = $null }   # original gone — fall through to a fresh message
    }
    if (-not $item) {
        $item = $ol.CreateItem(0)
        $item.Subject = $spec.subject
        $item.Body = $spec.body
    }
    if ($spec.to) { $item.To = $spec.to }
    if ($spec.cc) { $item.CC = $spec.cc }
    if ($spec.subject -and $spec.reply_entry_id -and $item.Subject -notlike "*$($spec.subject)*") {
        # keep the reply's RE: subject unless the caller explicitly overrode it
    }

    if ($spec.mode -eq 'draft') {
        $item.Save()
        Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $true; mode = 'draft'; entry_id = $item.EntryID; to = $item.To; subject = $item.Subject })
    }
    else {
        $to = $item.To; $subject = $item.Subject
        $item.Send()
        Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $true; mode = 'send'; to = $to; subject = $subject })
    }
}
catch {
    Write-Output (ConvertTo-Json -Compress -InputObject @{ ok = $false; error = $_.Exception.Message })
}
