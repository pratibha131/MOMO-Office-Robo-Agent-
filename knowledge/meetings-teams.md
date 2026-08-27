# Meetings & Teams (classic Outlook COM + Exchange)

Meetings are created in Outlook via COM. Because the mailbox is Exchange Online, the
meeting lands on BOTH the Outlook calendar and the Teams calendar automatically (they are
the same Exchange calendar). Invites are emailed to attendees by Exchange.

## Teams link
One-time setup (tell the user if links are missing): Outlook Web (outlook.office.com)
-> Settings -> Calendar -> Events and invitations -> enable "Add online meeting to all
meetings". With that on, Exchange auto-attaches a Teams link to every invite --
including ones Momo creates via COM.

## 1. Read the calendar (today / next N hours) — needed for briefs, heartbeat, conflicts
```powershell
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$cal = $ns.GetDefaultFolder(9)   # olFolderCalendar
$items = $cal.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]")
$start = Get-Date; $end = $start.AddHours(24)
$flt = "[Start] >= '" + $start.ToString("g") + "' AND [Start] <= '" + $end.ToString("g") + "'"
$res = $items.Restrict($flt)
$out = @()
foreach ($a in $res) {
  $out += [pscustomobject]@{ Subject=$a.Subject; Start=$a.Start.ToString("s"); End=$a.End.ToString("s");
    Organizer=$a.Organizer; Required=$a.RequiredAttendees; Optional=$a.OptionalAttendees; Location=$a.Location }
}
$out | ConvertTo-Json
```
IMPORTANT: with IncludeRecurrences, ALWAYS Sort then Restrict on a bounded window —
never enumerate all items.

## 2. Create a Teams meeting (draft -> review -> send, same rule as email)
```powershell
$appt = $ol.CreateItem(1)          # olAppointmentItem
$appt.MeetingStatus = 1            # olMeeting -> makes it an invite
$appt.Subject = "Project BCP — Sync"
$appt.Start = Get-Date "2026-08-28 15:00"
$appt.Duration = 30
$appt.Body = "Agenda: ..."
$r = $appt.Recipients.Add("person@philips.com"); $r.Type = 1   # 1=Required, 2=Optional
$null = $appt.Recipients.ResolveAll()
$appt.Save()
"ENTRYID=" + $appt.EntryID
# $appt.Display()  -> for review;   $appt.Send() -> ONLY after explicit yes
```
Store pending meeting EntryID in data/state.json like drafts.
Time parsing: user speaks casually ("tomorrow at 3", "Friday morning"). Compute concrete
datetimes; default duration 30 min; if no time given, propose the first free slot and ask.
Conflict check: read the calendar for that window first (section 1); warn if it clashes.

## 3. After sending
Say e.g. "Done — the invite with the Teams link is on both your calendars, and
Pratibha got it in her inbox." (Exchange emails the invite; no separate mail needed.
Only send a separate Outlook mail with the join link if the user explicitly asks.)

## Graph API fallback (optional, only if COM ever fails)
momo.mcp.json in the project root has a disabled Microsoft 365 MCP server entry
(Graph device-code login). Only suggest it if classic Outlook COM is unavailable.
