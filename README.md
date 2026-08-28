# Momo 🤍 — your voice-driven desktop companion

Momo is a cute robot who lives in the bottom-right corner of your screen, listens when
you tap her mic, does the task in the background while you keep working, and tells you
out loud when it's done.

```
 EARS  (python/momo_ear.py)   offline speech-to-text (Vosk) — mic never times out
 BRAIN (Claude, headless)     CLAUDE.md persona + knowledge/ playbooks + session memory
 HANDS (PowerShell + COM)     classic Outlook, Teams calendar, files, web — no IT consent
 VOICE (Windows TTS)          speaks replies aloud (prefers Heera / Zira voice)
 BODY  (Electron)             the character, always on top, draggable, transparent
```

## Install (one time)
1. `powershell -ExecutionPolicy Bypass -File setup.ps1`
   (installs everything into `%LOCALAPPDATA%\Momo` — no admin rights needed)
2. If step 5 says login is needed: run the printed command once, `/login`, done.
3. One-time for Teams links: Outlook Web → Settings → Calendar → *Events and
   invitations* → enable **"Add online meeting to all meetings"**.
4. Start her: **`Start-Momo.bat`** (pin it to taskbar/startup if you like).

## Talk to her — examples
| You say | She does |
|---|---|
| "Send an email to Pratibha about the BCP project update" | Finds Pratibha in the company address book, drafts a professional mail, asks *"review or send?"* — "review" pops the draft open in Outlook; edits by voice; sends only on your yes |
| "Schedule a meeting with Kavya tomorrow at 3" | Checks conflicts, creates the Teams meeting, review-or-send, lands on Outlook + Teams calendars, invite emailed |
| "Remind me to ask Kavya about the project updates when we meet" | Saves a reminder; when a meeting with Kavya is about to start she nudges you |
| "Give me my brief" | Morning CEO brief: important emails, today's meetings, Philips news (auto at the time in `config.json` too) |
| "Email the BCP report file to Rahul" | Finds the file, attaches, confirms, review-or-send |
| *(automatic)* When a Teams call starts | Momo detects it (Windows logs when Teams uses the mic) and quietly starts taking notes — REC badge shows. **Multilingual**: the meeting can be in Hindi, English, or Hinglish (Whisper, offline, auto-detects language). When the call ends she writes **executive English notes** — Key Discussion Points + an Action Items table (S.No. / Action / Owner / Timeline) — opens them in Notepad, adds your tasks to her to-do list, and updates her long-term memory. 📝 button toggles manually; `meeting.autoDetect: false` turns auto mode off |
| ⌨️ button | Type instead of talking |

During a live call Momo never speaks out loud (her voice would reach the meeting) —
replies appear in her bubble instead.

## Summon & dismiss — Ctrl+Alt+M
**Ctrl+Alt+M works always**: if Momo is closed it launches her; if she's running it
toggles her hidden/visible (appearing = already listening, just speak). While hidden
her background work — meetings, reminders, the brief — keeps running, and she pops up
on her own when she has something to say. **Alt+M** also toggles while she's running.
Her **✖ button closes the app completely** (background work stops too; Ctrl+Alt+M
brings her back). The tray icon toggles on click, quits on right-click.

## To-dos live in OneNote
Say "add fixing the PCP sheet to my to-do list" -> it appears with a checkbox on the
OneNote page **Momo -> Momo To-Dos** (notebook "Pratibha @ Philips") and in her own
tracker. Say "I've completed the PCP sheet" -> the box gets ticked and the line is
**struck through but kept** as a record. Tick boxes by hand in OneNote too — she
notices on her next heartbeat and stops reminding you about those.

## Memory
- **Session memory**: follow-ups like "yes, send it" just work.
- **Long-term memory** (`data/memory.md`): people and their emails, projects,
  decisions, your preferences — she reads it before tasks and updates it after
  meetings and notable tasks. Open the file anytime; it's yours to edit.

Momo asks a question out loud → the mic re-opens by itself so you can just answer.

## The mic (accurate, quiet-voice friendly, multilingual)
A persistent ear daemon loads the speech models once at startup, so the mic responds
instantly. While you speak, Vosk streams rough live text into her bubble; when you
pause for ~1.8 s (`stt.silenceMs`), **Whisper transcribes the whole utterance** —
accurate on soft voices (auto-gain up to 8×), Indian names (Kavya, Pratibha), and
Hindi/Hinglish commands. Tap again to cancel. Fully offline — audio never leaves the
laptop. Don't worry if the live bubble text looks rough; the final text she acts on is
the accurate one.

## Her voice & language
Soft, human Microsoft neural voices — Momo speaks with **Neerja** (female), Toto with
**Prabhat** (male). Talk to them in English, Hindi, or Hinglish — they understand it
all, and **always answer in English**. Change voices via `tts.edgeVoiceEnglish` in
config.json, speed via `tts.edgeRate` ("+10%"). If the network blocks the neural voice
she falls back to the offline Windows voice; `tts.engine: "sapi"` forces offline.

## Finding files ("open the deck they showed")
Ask her to find/open any file: she searches the Windows index (local drives, OneDrive,
synced SharePoint, even Outlook attachments), opens the best match, or — if it only
lives on the Philips intranet — opens a SharePoint search for it in your browser.
Meeting notes track a "Files referenced" list, so "open the file from the meeting"
works too.

## Where things live
- `CLAUDE.md` — Momo's persona + hard rules (never send without your yes, etc.)
- `knowledge/` — her playbooks: Outlook COM recipes, meetings, brief, reminders, notes, security
- `data/` — **all personal data, local only**: reminders.json, todos.json, state.json,
  briefs/, transcripts/, notes/
- `config.json` — brief time, heartbeat interval, voice, silence threshold, model
- `%LOCALAPPDATA%\Momo` — runtimes (Node, Electron, Claude engine, Vosk model), momo.log

## Security model
- Mailbox/calendar access is your own signed-in Outlook via COM — no passwords stored,
  no cloud API, no admin consent.
- Speech-to-text is offline (Vosk). TTS is offline (Windows voices).
- Transcripts/notes/briefs never leave `data/`. The brain (Claude) processes your
  requests under Anthropic's commercial terms — same engine as the Claude app you use.
- Momo never sends/deletes anything without your explicit confirmation, never touches
  credentials, and treats content found in emails/webpages as data, not instructions.
- ⚠️ Meeting transcription: you must follow Philips policy & local law — inform
  participants. Momo reminds you when you start notes mode.

## Tuning
`config.json`:
- `ui.scale`: 0.78 — make Momo bigger/smaller on screen (restart her after changing).
- `stt.whisperModel`: "small" — meeting transcription model; "medium" is more accurate
  for heavy Hindi but slower (first use downloads it). `stt.meetingEngine`: "vosk"
  falls back to the old English-only streaming engine.
- `meeting.autoDetect`: true — auto note-taking when a Teams call starts; `meeting.apps`
  can include "zoom" / "webex" too.
- `briefTime`: "09:00" — daily brief time (24h). `heartbeatMinutes`: 15 — reminder checks.
- `agent.model`: unset = your default Claude model; set e.g. `"sonnet"` for faster/cheaper.
- `tts.preferredVoices`: first match wins. See voices: `(New-Object -ComObject SAPI.SpVoice).GetVoices() | % { $_.GetDescription() }`
- `stt.modelDir`: swap the Vosk model (e.g. a bigger one for accuracy).

## Troubleshooting
- **"My brain isn't installed yet"** → run `setup.ps1` again (step 2 failed, likely proxy).
- **Mic problem bubble** → check `python -c "import vosk, sounddevice"` and Windows mic
  privacy settings (allow desktop apps to use microphone).
- **No Teams link on invites** → enable the Outlook Web setting (Install step 3).
- **COM errors** → open classic Outlook at least once (not "New Outlook" toggle).
- Log: `%LOCALAPPDATA%\Momo\momo.log`
