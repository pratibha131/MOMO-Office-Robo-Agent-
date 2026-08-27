# Daily CEO Brief  (prompt starts with [BRIEF])

Goal: a crisp spoken morning brief — like a chief of staff. Three sections, then save.

## 1. Important emails (since yesterday evening)
Via Outlook COM (see outlook-email.md): Inbox items from the last ~16 hours.
Rank importance: marked High Importance; sender is a frequent correspondent (appears in
Sent Items ranking); subject keywords (urgent, action, review, deadline, escalation,
approval); addressed directly vs CC. Pick the top 3-5. Ignore newsletters/automated mail.

## 2. Today's meetings
Calendar for today (meetings-teams.md section 1): time, subject, organizer. Mention gaps
or back-to-back stretches. Cross-check data/reminders.json and data/todos.json — if a
meeting matches a reminder person, or a task is due today, say so.

## 3. Philips business news
WebSearch: "Philips news" and "Royal Philips announcement" (last 24-48h). Health-tech
context only if it directly affects Philips. 2-3 headlines max, one line each. If nothing
noteworthy, say so in one line. Never invent news; only report actual search results.

## Output
- SPOKEN (final reply): about 30-60 seconds of speech. "Good morning! Quick brief: three
  emails need you — ... You have four meetings — ... And in Philips news — ..."
- SAVED: full detail to data/briefs/YYYY-MM-DD.md (markdown, with sources for news).
  Mention the saved file only if asked.
