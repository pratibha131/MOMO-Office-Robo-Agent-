// Momo — main process: window, ears (python STT), brain (claude headless), voice (SAPI TTS)
const { app, BrowserWindow, ipcMain, screen } = require('electron');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PROJECT = path.resolve(__dirname, '..');
const MOMO_HOME = path.join(process.env.LOCALAPPDATA || os.homedir(), 'Momo');
const NODE_DIR = path.join(MOMO_HOME, 'runtime', 'node');
const CLAUDE_EXE = path.join(MOMO_HOME, 'app', 'node_modules', '@anthropic-ai', 'claude-code', 'bin', 'claude.exe');
const DATA = path.join(PROJECT, 'data');

// ---------- config ----------
function loadConfig() {
  const cfg = JSON.parse(fs.readFileSync(path.join(PROJECT, 'config.json'), 'utf8'));
  cfg.stt.modelDir = cfg.stt.modelDir.replace(/%LOCALAPPDATA%/gi, process.env.LOCALAPPDATA || '');
  cfg.ui = cfg.ui || {};
  cfg.meeting = cfg.meeting || {};
  return cfg;
}
let config = loadConfig();

// ---------- state ----------
let win = null;
let earProc = null;          // command-mode listener
let meetingProc = null;      // meeting-mode listener
let meetingTranscript = null;
let meetingAutoStarted = false;   // was capture started by auto-detection?
let meetingStopping = false;      // stopMeeting() initiated this shutdown
let meetingSuppressed = false;    // user stopped capture during an ongoing call
let meetingFreeChecks = 0;
let suppressFreeChecks = 0;
let micCheckInFlight = false;
let lastMicInUse = false;
let agentBusy = false;
let pendingJobs = [];             // queued runAgent jobs (never drop meeting notes)
let ttsProc = null;
let sessionId = null;
let sessionLast = 0;
let lastBriefDate = null;
let quitting = false;
let state = 'idle';

function send(ch, payload) { if (win && !win.isDestroyed()) win.webContents.send(ch, payload); }
function setState(s) { state = s; send('momo:state', s); }
function bubble(text, sticky) { send('momo:bubble', { text, sticky: !!sticky }); }
function log(...a) { console.log(new Date().toISOString(), ...a); }

// ---------- window ----------
function createWindow() {
  const { workArea } = screen.getPrimaryDisplay();
  const scale = Number(config.ui.scale) || 1;
  const W = Math.round(230 * scale), H = Math.round(360 * scale);
  win = new BrowserWindow({
    width: W, height: H,
    x: workArea.x + workArea.width - W - 10,
    y: workArea.y + workArea.height - H - 6,
    transparent: true, frame: false, resizable: false,
    alwaysOnTop: true, skipTaskbar: true, hasShadow: false,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true }
  });
  win.setAlwaysOnTop(true, 'screen-saver');
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.webContents.on('did-finish-load', () => {
    win.webContents.setZoomFactor(scale);   // always set: zoom persists per-host otherwise
  });
  win.once('ready-to-show', () => {
    setTimeout(() => speak("Hi, I'm Momo! Tap the mic whenever you need me."), 800);
  });
}

// ---------- TTS (Windows SAPI via PowerShell; prefers en-IN female voice) ----------
function speak(text, onDone) {
  const clean = String(text)
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/[*_#`>|]/g, ' ')
    .replace(/\[[^\]]*\]\([^)]*\)/g, ' ')
    .replace(/https?:\S+/g, ' a link ')
    .replace(/\s+/g, ' ').trim();
  if (!clean) { if (onDone) onDone(); return; }
  // Stay silent out loud when her voice could reach a call (capture running, or the
  // mic is in use by a meeting app) or be picked up by her own open ear.
  if (meetingProc || lastMicInUse || earProc) {
    bubble(clean, true);
    if (meetingProc) setState('meeting');
    else if (state === 'thinking' || state === 'speaking') setState('idle');
    if (onDone) onDone();
    return;
  }
  stopSpeaking();
  setState('speaking');
  bubble(clean, true);
  const b64 = Buffer.from(clean, 'utf8').toString('base64');
  const prefs = (config.tts.preferredVoices || []).map(v => "'" + v + "'").join(',');
  const ps = [
    'Add-Type -AssemblyName System.Speech;',
    '$s = New-Object System.Speech.Synthesis.SpeechSynthesizer;',
    '$s.Rate = ' + (Number(config.tts.rate) || 0) + ';',
    '$names = $s.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name };',
    'foreach ($p in @(' + prefs + ')) { $m = $names | Where-Object { $_ -like ("*" + $p + "*") } | Select-Object -First 1; if ($m) { $s.SelectVoice($m); break } }',
    "$t = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + b64 + "'));",
    '$s.Speak($t);'
  ].join(' ');
  const p = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps], { windowsHide: true });
  ttsProc = p;
  const finish = () => {
    if (ttsProc !== p) return;          // a newer utterance took over
    ttsProc = null;
    if (state === 'speaking') setState(meetingProc ? 'meeting' : 'idle');
    if (onDone) onDone();
  };
  p.on('exit', finish);
  p.on('error', e => { log('tts error', e.message); finish(); });
}
function stopSpeaking() {
  const p = ttsProc;
  ttsProc = null;                        // null first: the old exit handler must not fire onDone
  if (p) { try { p.kill(); } catch (e) {} }
}

// ---------- Ears (python STT) ----------
function pyEnv() {
  return Object.assign({}, process.env, { PYTHONUTF8: '1' });
}
function pyArgs(extra) {
  return [path.join(PROJECT, 'python', 'momo_ear.py'),
    '--model', config.stt.modelDir, '--silence-ms', String(config.stt.silenceMs)].concat(extra);
}
function startEar() {
  if (earProc) return;
  stopSpeaking();
  // NOTE: starting the command ear DURING meeting capture is allowed on purpose —
  // you can ask Momo something mid-call; Windows shares the mic between captures.
  const p = spawn('python', pyArgs([]), { cwd: PROJECT, env: pyEnv(), windowsHide: true });
  earProc = p;
  setState('listening');
  bubble('Listening…');
  let buf = '';
  p.stdout.on('data', d => {
    if (earProc !== p) return;          // stale listener: ignore leftovers
    buf += d.toString();
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line) continue;
      let msg; try { msg = JSON.parse(line); } catch (e) { continue; }
      if (msg.event === 'partial') send('momo:partial', msg.text);
      else if (msg.event === 'final' && msg.text) { stopEar(); handleCommand(msg.text); return; }
      else if (msg.event === 'error') {
        stopEar();
        setState(meetingProc ? 'meeting' : 'idle');
        bubble('Mic problem: ' + msg.message, true);
        log('ear error', msg.message);
        return;
      }
    }
  });
  p.stderr.on('data', d => log('ear:', d.toString().trim()));
  p.on('exit', () => {
    if (earProc !== p) return;
    earProc = null;
    if (state === 'listening') setState(meetingProc ? 'meeting' : 'idle');
  });
  p.on('error', e => {
    if (earProc !== p) return;
    earProc = null;
    setState(meetingProc ? 'meeting' : 'idle');
    bubble('Mic engine failed to start: ' + e.message, true);
  });
}
function stopEar() {
  const p = earProc;
  earProc = null;
  if (p) { try { p.kill(); } catch (e) {} }
}

// ---------- Meeting capture ----------
function startMeeting(auto) {
  if (meetingProc) return;
  stopSpeaking();          // never let an in-flight utterance play into the call
  stopEar();
  const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
  meetingTranscript = path.join(DATA, 'transcripts', 'meeting-' + stamp + '.txt');
  const extra = ['--meeting', '--transcript', meetingTranscript,
    '--engine', config.stt.meetingEngine || 'whisper',
    '--whisper-model', config.stt.whisperModel || 'small',
    '--chunk-sec', String(config.stt.chunkSec || 45)];
  const p = spawn('python', pyArgs(extra), { cwd: PROJECT, env: pyEnv(), windowsHide: true });
  meetingProc = p;
  p.stderr.on('data', d => log('meeting-ear:', d.toString().trim()));
  let mbuf = '';
  p.stdout.on('data', d => {
    if (meetingProc !== p) return;
    mbuf += d.toString();
    let i;
    while ((i = mbuf.indexOf('\n')) >= 0) {
      const line = mbuf.slice(0, i).trim(); mbuf = mbuf.slice(i + 1);
      if (!line) continue;
      log('meeting-ear:', line);
      let msg; try { msg = JSON.parse(line); } catch (e) { continue; }
      if (msg.event === 'error') bubble('Note-taking problem: ' + msg.message, true);
    }
  });
  let ended = false;
  const onEnd = (spawnFailMsg) => {
    if (ended) return; ended = true;
    if (meetingProc === p) meetingProc = null;
    meetingAutoStarted = false;
    const expected = meetingStopping;
    meetingStopping = false;
    send('momo:rec', false);
    if (quitting) return;
    if (spawnFailMsg) {                 // never even started: no notes to write
      setState(earProc ? state : 'idle');
      bubble(spawnFailMsg, true);
      return;
    }
    if (!expected) bubble('Meeting capture stopped unexpectedly — writing up what I have.', true);
    finishMeetingNotes();
  };
  p.on('exit', () => onEnd());
  p.on('error', e => { log('meeting-ear spawn error', e.message); onEnd('Could not start the note-taker (is python installed?).'); });
  meetingAutoStarted = !!auto;
  setState('meeting');
  send('momo:rec', true);
  if (auto) {
    bubble('Meeting detected — taking notes quietly.', true);   // no voice: she is silent during calls
  } else {
    speak('Taking meeting notes from now.');
  }
}
function stopMeeting() {
  if (!meetingProc || meetingStopping) return;
  meetingStopping = true;
  setState('thinking');
  bubble('Meeting ended — finishing the transcript…', true);
  const p = meetingProc;
  // graceful stop: the ear transcribes all buffered audio and writes an end marker
  // (whisper may need a couple of minutes for the final chunks)
  try { p.stdin.write('stop\n'); } catch (e) { try { p.kill(); } catch (e2) {} }
  setTimeout(() => { try { p.kill(); } catch (e) {} }, 300000);   // fallback
}
function finishMeetingNotes() {
  const t = meetingTranscript;
  meetingTranscript = null;
  const empty = !t || !fs.existsSync(t) || fs.statSync(t).size < 200;
  if (empty) {
    if (!earProc) setState('idle');
    bubble('Meeting ended — nothing much was said, so no notes this time.');
    return;
  }
  setState('thinking');
  bubble('Writing up your meeting notes…', true);
  // speak() itself decides voice vs bubble (live call, open ear) and restores state
  runAgent('[MEETING-NOTES] The meeting just ended. Transcript file: ' + t + '. Follow knowledge/meeting-notes.md.',
    { fresh: true, queue: true }, reply => speak(reply));
}

// ---------- Auto meeting detection (Windows tracks when Teams uses the mic) ----------
function checkMeetingMic(cb) {
  const rx = (config.meeting.apps || ['teams']).join('|');
  const ps = [
    "$rx = '" + rx + "';",
    "$root = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\microphone';",
    '$inuse = $false;',
    'Get-ChildItem $root -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match $rx } | ForEach-Object {',
    '  $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue;',
    '  if ($p.LastUsedTimeStart -and (-not $p.LastUsedTimeStop -or $p.LastUsedTimeStart -gt $p.LastUsedTimeStop)) { $inuse = $true }',
    '};',
    "if ($inuse) { 'IN_USE' } else { 'FREE' }"
  ].join(' ');
  const proc = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps], { windowsHide: true });
  let out = '';
  let settled = false;
  const settle = v => { if (!settled) { settled = true; cb(v); } };
  proc.stdout.on('data', d => out += d);
  proc.on('exit', () => settle(out.indexOf('IN_USE') >= 0));
  proc.on('error', () => settle(false));
}
function meetingWatchTick() {
  if (config.meeting.autoDetect === false) return;
  if (micCheckInFlight) return;
  micCheckInFlight = true;
  checkMeetingMic(inUse => {
    micCheckInFlight = false;
    lastMicInUse = inUse;
    if (inUse) {
      meetingFreeChecks = 0;
      suppressFreeChecks = 0;
      if (!meetingProc && !meetingStopping && !meetingSuppressed) startMeeting(true);
    } else {
      // re-arm suppression only after the call has been over for a while (readings flap)
      if (meetingSuppressed) {
        suppressFreeChecks++;
        if (suppressFreeChecks >= (config.meeting.endAfterChecks || 2)) {
          meetingSuppressed = false;
          suppressFreeChecks = 0;
        }
      }
      if (meetingProc && meetingAutoStarted && !meetingStopping) {
        meetingFreeChecks++;
        if (meetingFreeChecks >= (config.meeting.endAfterChecks || 2)) {
          meetingFreeChecks = 0;
          stopMeeting();
        }
      } else {
        meetingFreeChecks = 0;
      }
    }
  });
}

// ---------- Brain (claude headless) ----------
function drainJobs() {
  if (!agentBusy && pendingJobs.length) {
    const j = pendingJobs.shift();
    runAgent(j.prompt, j.opts, j.cb);
  }
}
function runAgent(prompt, opts, cb) {
  opts = opts || {};
  if (agentBusy) {
    if (opts.queue) {
      pendingJobs.push({ prompt, opts, cb });
      if (!opts.background) bubble('On it — right after I finish the current task…', true);
      return;
    }
    if (!opts.background) bubble('One moment — still finishing the last task…', true);
    return;
  }
  if (!fs.existsSync(CLAUDE_EXE)) {
    pendingJobs = [];
    speak('My brain is not installed yet. Please run the setup script first.');
    return;
  }
  agentBusy = true;
  if (!opts.background) setState('thinking');

  const idleMs = (config.agent.sessionIdleResetMinutes || 30) * 60000;
  const resume = !opts.fresh && sessionId && (Date.now() - sessionLast) < idleMs;
  const args = ['-p', prompt, '--output-format', 'json', '--dangerously-skip-permissions'];
  if (config.agent.model) args.push('--model', config.agent.model);
  if (resume) args.push('--resume', sessionId);
  if (config.agent.useMcp) args.push('--mcp-config', path.join(PROJECT, 'momo.mcp.json'));

  const env = Object.assign({}, process.env, { PATH: NODE_DIR + ';' + process.env.PATH });
  const proc = spawn(CLAUDE_EXE, args, { cwd: PROJECT, env, windowsHide: true });
  let out = '', err = '';
  const timeout = setTimeout(() => { try { proc.kill(); } catch (e) {} }, (opts.background ? 240 : 420) * 1000);
  let settled = false;
  const settle = fn => {
    if (settled) return; settled = true;
    clearTimeout(timeout);
    agentBusy = false;
    try { fn(); } catch (e) { log('agent callback error', e); }
    drainJobs();
  };
  proc.stdout.on('data', d => out += d);
  proc.stderr.on('data', d => err += d);
  proc.on('error', e => settle(() => {
    log('agent spawn error', e.message);
    if (!opts.background) speak('Sorry, I could not start my brain process. Check my log.');
  }));
  proc.on('exit', code => settle(() => {
    let reply = null;
    try {
      const j = JSON.parse(out);
      reply = j.result ? String(j.result).trim() : null;   // empty result = failure
      if (!opts.background) { sessionId = j.session_id || sessionId; sessionLast = Date.now(); }
    } catch (e) {
      log('agent parse fail. code=', code, 'err=', err.slice(0, 500), 'out=', out.slice(0, 500));
    }
    if (!reply) {
      // speak() manages state itself; background failures stay silent and untouched
      if (!opts.background) speak('Sorry, I hit a problem with that one. Check my log for details.');
      return;
    }
    cb(reply);
  }));
}

function handleCommand(text) {
  send('momo:partial', '');
  bubble('“' + text + '”', true);
  setState('thinking');
  // queue:true — a command spoken while another task runs is done next, never dropped
  runAgent(text, { queue: true }, reply => {
    speak(reply, () => {
      // If Momo asked a question, reopen the mic so the user can just answer.
      if (/\?\s*$/.test(reply) && !meetingProc) startEar();
    });
  });
}

// ---------- Schedulers ----------
function heartbeatTick() {
  if (agentBusy || meetingProc || earProc || state === 'speaking') return;
  if (!config.heartbeatMinutes) return;
  runAgent('[HEARTBEAT] Time: ' + new Date().toString() + '. Follow the heartbeat procedure in knowledge/reminders-todos.md. Reply SILENT if nothing is due.',
    { fresh: true, background: true }, reply => {
      if (reply && reply.toUpperCase() !== 'SILENT') speak(reply);
    });
}
function briefTick() {
  const now = new Date();
  const hhmm = now.toTimeString().slice(0, 5);
  const today = now.toDateString();
  if (hhmm === config.briefTime && lastBriefDate !== today &&
      !agentBusy && !meetingProc && !earProc && state !== 'speaking') {
    lastBriefDate = today;
    setState('thinking');
    bubble('Preparing your morning brief…', true);
    runAgent('[BRIEF] Morning brief for ' + now.toDateString() + '. Follow knowledge/daily-brief.md.',
      { fresh: true }, reply => speak(reply));
  }
}

// ---------- IPC ----------
ipcMain.on('mic-toggle', () => {
  if (ttsProc) { stopSpeaking(); setState(meetingProc ? 'meeting' : 'idle'); return; }
  if (earProc) { stopEar(); setState(meetingProc ? 'meeting' : 'idle'); bubble('Okay, cancelled.'); }
  else startEar();
});
ipcMain.on('meeting-toggle', () => {
  if (meetingProc) {
    meetingSuppressed = true;      // any manual stop: don't auto-restart during this call
    suppressFreeChecks = 0;
    stopMeeting();
  } else {
    startMeeting(false);
  }
});
ipcMain.on('text-command', (e, text) => { if (text && text.trim()) handleCommand(text.trim()); });
ipcMain.on('momo-quit', () => { app.quit(); });

// ---------- app lifecycle ----------
if (!app.requestSingleInstanceLock()) app.quit();
app.whenReady().then(() => {
  createWindow();
  setInterval(heartbeatTick, Math.max(5, config.heartbeatMinutes || 15) * 60000);
  setInterval(briefTick, 30000);
  setInterval(meetingWatchTick, Math.max(5, config.meeting.checkSeconds || 15) * 1000);
});
app.on('before-quit', () => {
  quitting = true;
  stopEar();
  if (meetingProc) { try { meetingProc.stdin.write('stop\n'); } catch (e) {} try { meetingProc.kill(); } catch (e) {} }
  stopSpeaking();
});
app.on('window-all-closed', () => app.quit());
