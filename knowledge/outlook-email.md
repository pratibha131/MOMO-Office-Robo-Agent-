# Outlook email automation (classic Outlook, COM via PowerShell)

Outlook classic is installed and COM-registered. All operations use the signed-in
profile — no credentials, no Graph consent needed. Run everything via PowerShell.

## 0. Boilerplate (start of every Outlook script)
```powershell
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
```
If this throws "cannot create COM object", Outlook may be set to "New Outlook" — tell the
user to open classic Outlook once (toggle off "New Outlook" switch, top right).

## 1. Find a person's email (user only speaks a name)
Strategy, in order — combine results and rank:

**a) Global Address List (whole company):**
```powershell
$r = $ns.CreateRecipient("Pratibha")   # spoken name, partial ok
$null = $r.Resolve()
if ($r.Resolved) {
  $ae = $r.AddressEntry
  $smtp = $null
  if ($ae.Type -eq "EX") { $exu = $ae.GetExchangeUser(); if ($exu) { $smtp = $exu.PrimarySmtpAddress; $dept = $exu.Department; $title = $exu.JobTitle } }
  else { $smtp = $ae.Address }
  "$($ae.Name) <$smtp> $dept $title"
}
```
If not resolved (ambiguous), search the GAL address list:
```powershell
$gal = $ns.GetGlobalAddressList()
$hits = @()
foreach ($e in $gal.AddressEntries) {
  if ($e.Name -like "*pratibha*") {
    $exu = $e.GetExchangeUser()
    if ($exu -and $exu.PrimarySmtpAddress) { $hits += [pscustomobject]@{Name=$e.Name; Mail=$exu.PrimarySmtpAddress; Dept=$exu.Department} }
    if ($hits.Count -ge 8) { break }
  }
}
$hits | ConvertTo-Json
```
NOTE: iterating the full GAL is slow at large companies — `CreateRecipient/Resolve` first;
only fall back to iteration when unresolved, and break early.

**b) People the user actually mails (disambiguation/ranking):** scan Sent Items:
```powershell
$sent = $ns.GetDefaultFolder(5)  # olFolderSentMail
$items = $sent.Items; $items.Sort("[SentOn]", $true)
$seen = @{}
for ($i=1; $i -le [Math]::Min(300, $items.Count); $i++) {
  $m = $items[$i]; if ($m.MessageClass -notlike "IPM.Note*") { continue }
  foreach ($rcp in $m.Recipients) {
    if ($rcp.Name -like "*pratibha*") { $addr = $rcp.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E"); $seen[$addr] = 1 + ($seen[$addr] | ForEach-Object {$_}) }
  }
}
$seen | ConvertTo-Json
```
Prefer a GAL match the user has mailed before. If 2+ plausible people remain, ASK the
user (speak name + department for each).

## 2. Search mail for context (recent thread on a topic / from a person)
```powershell
$inbox = $ns.GetDefaultFolder(6)   # olFolderInbox
$flt = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%BCP%'"
$res = $inbox.Items.Restrict($flt); $res.Sort("[ReceivedTime]", $true)
for ($i=1; $i -le [Math]::Min(5,$res.Count); $i++) { $m=$res[$i]; "$($m.ReceivedTime) | $($m.SenderName) | $($m.Subject)" }
```
Received in a date window: `$inbox.Items.Restrict("[ReceivedTime] >= '" + (Get-Date).AddDays(-1).ToString("g") + "'")`
Read a body: `$m.Body` (plain) — take only first ~2000 chars for context.

## 3. Compose a professional draft (never .Send here)
```powershell
$mail = $ol.CreateItem(0)
$mail.To = "person@philips.com"
$mail.Subject = "Project BCP — Status Update"
$sigPath = Get-ChildItem "$env:APPDATA\Microsoft\Signatures\*.htm" -ErrorAction SilentlyContinue | Select-Object -First 1
$sig = ""; if ($sigPath) { $sig = Get-Content $sigPath.FullName -Raw }
$body = @"
<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f1f1f'>
<p>Dear Pratibha,</p>
<p>...clear, courteous, corporate body — short paragraphs...</p>
<p>Best regards,</p>
</div>
"@
$mail.HTMLBody = $body + $sig
$mail.Save()
"ENTRYID=" + $mail.EntryID
```
Style: formal-friendly corporate tone; greeting by first name; 2–4 short paragraphs;
explicit ask/action; courteous close. Use context from recent thread when relevant.
**Save the EntryID into data/state.json** as `pending_draft` {entryid, to, subject}.

## 4. Review / edit / send the pending draft
```powershell
$mail = $ns.GetItemFromID("ENTRYID_HERE")
$mail.Display()          # REVIEW: pops the draft window open on screen
# edits: modify $mail.HTMLBody / $mail.Subject / $mail.To then $mail.Save()
$mail.Send()             # ONLY after the user explicitly says yes
```
After sending, clear pending_draft from data/state.json and confirm aloud.

## 5. Attachments (file transfer by email)
`$mail.Attachments.Add("C:\full\path\file.xlsx")` — find the file first with Glob/dir if
the user only gave a rough name; confirm the exact file aloud before sending.

## Gotchas
- Always index Items collections from 1, not 0.
- Wrap COM scripts so they print machine-readable output (ConvertTo-Json) — parse it.
- One PowerShell process per script; COM objects don't persist between calls.
- If Outlook shows a security prompt on .Send (rare with signed-in Outlook), tell the user.
