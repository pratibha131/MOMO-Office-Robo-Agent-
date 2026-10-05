# Meeting notes from transcripts  (prompt starts with [MEETING-NOTES])

When meeting capture ends, Momo receives the path to a transcript file in
data/transcripts/. Lines look like `[you +MM:SS] …` (the user speaking) and
`[them +MM:SS] …` (other participants via system audio). The transcript is imperfect
machine transcription and may be in ANY language or mix — Hindi, English, Hinglish.

## Language rule (non-negotiable)
The notes are ALWAYS written in polished, executive-grade ENGLISH, whatever language
the meeting was held in. Translate faithfully; keep proper nouns, project names, and
technical terms as spoken (BCP, VCP, EOS, KPI names, people's names).

## Output format (match this register exactly — concise, outcome-oriented, specific)
Write to data/notes/YYYY-MM-DD-<slug>.md:

```markdown
# <Meeting title — from calendar if a matching event exists, else best guess> — <date>
**Attendees:** <from the matching calendar event; else "(from audio)">

## Key Discussion Points
- <One bullet per topic. Lead with the subject, then the decision/status/implication.
  Style example: "VCP reporting format will be revised to a more dashboard-based view,
  with initiatives linked to specific KPIs and quarter-wise impact tracking rather
  than only year-wise reporting.">
- <8–12 bullets max. Merge fragments about the same topic into one strong bullet.
  No filler ("was discussed", "they talked about") — state WHAT was decided/agreed/
  flagged and WHY it matters.>

## Action Items

| S.No. | Action | Owner | Timeline |
|---|---|---|---|
| 1 | <Verb-first, specific, self-contained action> | <Name(s) as heard> | <as stated: "Next review", "WIP", "Before 15 Sept", "Ongoing"…> |

## Action Items for Pratibha
- <only the user's own tasks, restated crisply>

## Files referenced
- <every file/deck/sheet mentioned in the meeting: best-guess name + who shared it.
  Omit the section if none. The user may later say "open the file from the meeting"
  — knowledge/find-files.md uses this list to locate it.>
```

Rules of quality:
- Decisions > descriptions. If a number, date, metric, or owner was said — capture it.
- Uncertain attribution or garbled audio: append "(unclear)" rather than guessing.
- If the transcript has too little content for a section, keep the section with a
  single honest line ("No formal action items were assigned.").

## Then
1. Open the notes in Notepad: Start-Process notepad.exe with the absolute path.
2. Add every task assigned to the USER to data/todos.json (reminders-todos.md schema)
   AND as checkbox lines on the OneNote "Momo To-Dos" page (knowledge/onenote-todos.md).
   If a deadline was "before our next meeting", use due_type
   "before_next_meeting_with:<assigner>".
3. Update long-term memory (data/memory.md): under the right project heading, add
   1–3 lines of durable facts from this meeting (decisions made, who owns what,
   changed deadlines). Keep memory.md tidy — merge, don't just append duplicates.
4. Speak a 2-sentence summary: "Notes are open in Notepad. You picked up two tasks —
   the test report for Rahul by Tuesday, and the risk sheet before your next sync."

## MOM Excel — EVERY meeting gets its own workbook (in ADDITION to the .md notes)
For every meeting, generate a standalone MOM file with the ready-made tool (exact
VCP styling — Aptos Narrow, steel-blue title bar, yellow section headers, bold
decisions, bordered action table — is baked into it):
1. Write `mom.json`: {"title", "date" DD-MM-YYYY, "decisions": [...],
   "actions": [{"action","owner","timing"}]} (owner/timing "(to confirm)" when
   attribution is uncertain — never invent).
2. Run: `python python/momo_mom.py --json mom.json
   --out "data/notes/MOM - <Topic> - YYYY-MM-DD.xlsx"` and open the file
   (Start-Process) alongside the Notepad notes.
3. If it was a VCP review (subject/attendees mention VCP, organizer Aditya
   Jindal / VCP team), ALSO append a sheet in the central workbook
   `Desktop/VCP-NEW-FORMAT/MOM VCP Meetings.xlsx` with the same content
   (openpyxl, back it up first; if locked save a copy suffixed _with_new_MOM).

## Writing decisions at Copilot quality
Decisions/outcomes must read like commitments, not observations. Style rules
(learned from the user's preferred exemplar):
- Outcome-first phrasing: "X will shift to...", "A standardized framework will
  connect...", "The team aligned on..." — never "there was a discussion about".
- Merge fragments about one topic into ONE substantive line (aim 8-12 total).
- Carry every number, metric name (EBITDA, CLV, market share), tool, and deadline
  the transcript supports; generalize gracefully over garbled patches instead of
  quoting noise, and only mark "(unclear)" when genuinely undecipherable.

## Transcription-quality glossary (IMPORTANT — keep it fed)
`data/meeting_glossary.txt` primes the speech model during meetings and commands —
it is the difference between "काविय बारदवाज" and "Kavya Bhardwaj" in Hinglish audio.
MAINTAIN IT: before summarizing, and whenever you meet new people/projects, append
their names and recurring jargon (keep the file under ~120 lines, merge duplicates).
Heartbeat may also add upcoming-meeting attendee names from the calendar.

## Honesty about hard limits
Momo hears ONE mixed audio stream (mic + speakers) — unlike Teams Copilot, which
gets clean per-speaker cloud audio. So: attribute statements to named speakers only
when the content makes it obvious; otherwise use "the team" / "it was discussed".
Never invent owners for actions — write "(to confirm)" and tell the user which
ones need confirming.

## Privacy
Transcripts and notes NEVER leave data/. Do not email or upload them unless the user
explicitly asks, and remind them the file contains meeting content when they do.
