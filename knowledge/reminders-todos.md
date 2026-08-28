# Reminders, To-dos & the Heartbeat  (prompt starts with [HEARTBEAT])

Two JSON stores in data/ — always read-modify-write the whole file, keep valid JSON.

**OneNote mirror (IMPORTANT):** every to-do also lives on the user's OneNote page
"Momo To-Dos" (knowledge/onenote-todos.md has the tested recipes). Adding a todo →
add the OneNote checkbox line too. User says they completed something → set
done:true here AND tick + strike through the OneNote line (keep the struck line
visible as a record — never delete). If the two ever disagree, OneNote wins for
done-status (the user may tick boxes by hand); import manual OneNote ticks back
into todos.json when you notice them.

## data/reminders.json — user-requested reminders
```json
[{ "id": "r-20260827-1", "created": "2026-08-27", "text": "ask Kavya about project updates",
   "person": "Kavya", "trigger": { "type": "before_meeting_with", "minutes_before": 10 },
   "delivered": false }]
```
Trigger types:
- "before_meeting_with": fire when a calendar meeting starting within minutes_before has
  an attendee/organizer whose name fuzzy-matches person.
- "datetime": fire when now >= value (ISO datetime in a "value" field).
- "next_meeting_with": like before_meeting_with but only the FIRST upcoming match.

## data/todos.json — tasks (mostly harvested from meeting notes)
```json
[{ "id": "t-20260827-1", "task": "Share VCP test report", "assigned_by": "Rahul",
   "due": "2026-09-02", "due_type": "date", "source": "notes/2026-08-27-bcp-sync.md",
   "done": false, "reminded": false }]
```
due_type:
- "date": remind 1 day before due (and again on the day, if still not done).
- "before_next_meeting_with:<Name>": find the next calendar meeting with that person;
  remind 1 day before it. Re-evaluate each heartbeat — the meeting may move.

## Heartbeat procedure (runs every ~15 min; MUST be fast and quiet)
1. Read both JSON files. If both are empty -> reply SILENT (skip the calendar entirely).
2. Read the calendar for the next 24h once (meetings-teams.md section 1).
3. Reminders: for each undelivered reminder whose trigger fires now -> include it in the
   spoken output, phrased helpfully: "Heads up — Kavya is in your 3 PM meeting. That is
   who you wanted to ask about the project updates." Then set delivered:true.
4. Todos: BEFORE nudging, read the OneNote "Momo To-Dos" page once and import any
   manually ticked boxes into todos.json as done:true (the user ticks by hand too).
   Then: due tomorrow (or the before-next-meeting rule) and not done and not
   reminded -> "Reminder: your report for Rahul is due tomorrow, before your sync
   with him." Set reminded:true. Never nudge a task whose OneNote line is struck.
5. Recap offer: if a meeting starts within ~10 min and data/notes/ contains notes from a
   previous meeting with the same organizer or similar subject -> ask: "Your BCP sync
   with Rahul starts in ten minutes — want a quick recap of the last one?" (The yes/no
   arrives as the next voice command; on yes, summarize that notes file aloud in 4-6
   sentences.) Track offered recaps in data/state.json to avoid repeating the offer.
6. Nothing due -> reply exactly SILENT.

## Adding a reminder (user: "remind me to ask Kavya about updates in our meeting")
Parse person + text + trigger type, append to reminders.json, confirm in one short
sentence: "Got it — I will nudge you when you are about to meet Kavya."
