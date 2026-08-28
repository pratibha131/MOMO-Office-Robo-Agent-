# Finding & opening files (PC, SharePoint, Outlook attachments)

The user asks things like "open the deck they showed in the meeting", "find the BCP
excel", "open the file Rahul mentioned". Files may live locally, in synced
OneDrive/SharePoint folders, as Outlook attachments, or only on Philips SharePoint.

## 1. Windows Search index — ALWAYS try this first (fast, covers almost everything)
It indexes local drives, OneDrive, synced SharePoint libraries AND Outlook mail
attachments in one query (verified working on this machine):

```powershell
$conn = New-Object System.Data.OleDb.OleDbConnection "Provider=Search.CollatorDSO;Extended Properties='Application=Windows'"
$conn.Open()
$q = "SELECT TOP 15 System.ItemPathDisplay, System.DateModified, System.Size FROM SYSTEMINDEX " +
     "WHERE (System.FileName LIKE '%dashboard%' OR CONTAINS(System.FileName, '""dashboard""')) " +
     "AND System.FileExtension IN ('.pptx','.ppt','.xlsx','.xls','.docx','.pdf') " +
     "ORDER BY System.DateModified DESC"
$cmd = New-Object System.Data.OleDb.OleDbCommand($q, $conn)
$r = $cmd.ExecuteReader()
while ($r.Read()) { "$($r.GetValue(1)) | $($r.GetValue(0))" }
$r.Close(); $conn.Close()
```
- Search TOKENS from what the user said (file topic words), not exact phrases.
- Drop the extension filter if the type is unknown.
- Content search too: `CONTAINS(System.Search.Contents, '"project bcp"')` finds files
  BY WHAT'S INSIDE them when the name is unknown.

## 2. Interpret the results
- Normal path (C:\..., OneDrive...) → a real file. Open it:
  `Start-Process "C:\full\path\file.xlsx"`
- Path starting with `/Someone@philips.com/...` → an **Outlook attachment**, not a
  file on disk. Open it via Outlook COM — find the newest mail carrying it:
```powershell
$ol = New-Object -ComObject Outlook.Application; $ns = $ol.GetNamespace("MAPI")
foreach ($fid in 6, 5) {   # Inbox, then Sent Items
  $items = $ns.GetDefaultFolder($fid).Items
  $items.Sort("[ReceivedTime]", $true)
  $res = $items.Restrict("@SQL=""urn:schemas:httpmail:hasattachment"" = 1")
  for ($i = 1; $i -le [Math]::Min(150, $res.Count); $i++) {
    $m = $res[$i]
    foreach ($a in $m.Attachments) {
      if ($a.FileName -like "*dashboard*") {
        $p = Join-Path $env:TEMP $a.FileName
        $a.SaveAsFile($p); Start-Process $p
        "OPENED $($a.FileName) from '$($m.Subject)' ($($m.SenderName))"; return
      } } } }
```

## 3. Philips SharePoint (intranet) — when it's not on the machine
No API access is available, but the user's browser is SSO'd into SharePoint. Open a
Microsoft Search results page scoped to files — the user picks the file there:
```powershell
Start-Process "https://philips.sharepoint.com/_layouts/15/sharepoint.aspx?q=dashboard%20redesign"
```
Also useful: `https://www.office.com/search/files?q=<query>` (searches all M365).
Say what you did: "I could not find it on your machine, so I opened a SharePoint
search for it — the results are in your browser."

## 4. Choosing and confirming
- Rank: newest first; prefer matches on more of the user's words; prefer Downloads/
  Desktop/OneDrive over temp paths.
- ONE strong match → open it and confirm aloud ("Opening the MoS dashboard deck from
  the fourteenth of July").
- 2–3 plausible matches → ask, speaking name + date for each.
- Zero local matches → SharePoint search page (section 3).

## 5. Meeting connection
Transcripts in data/transcripts/ often carry the reference ("the excel Ketan shared",
"iss deck mein"). When the user says "the file from the meeting", read the latest
transcript/notes to recover the file's name or topic words, then search as above.
The notes template has a "Files referenced" section — check data/notes/ first.
