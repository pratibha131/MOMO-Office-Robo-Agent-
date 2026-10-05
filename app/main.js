// Momo — main process: window, ear daemon (python STT), brain (claude headless), voice (SAPI TTS)
const { app, BrowserWindow, ipcMain, screen, globalShortcut, Tray, Menu, nativeImage } = require('electron');
const { spawn, execSync } = require('child_process');
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

// ---------- profile (who the user is + which character runs) ----------
function loadProfile() {
  try { return JSON.parse(fs.readFileSync(path.join(DATA, 'profile.json'), 'utf8')); }
  catch (e) { return {}; }
}
const profile = loadProfile();
const CHAR = profile.character === 'toto' ? 'toto' : 'momo';
const CHAR_NAME = profile.assistantName || (CHAR === 'toto' ? 'Toto' : 'Momo');

// ---------- state ----------
let win = null;
let ear = null;               // persistent python listening daemon
let earReady = false;
let earRestarts = 0;
let listening = false;
let listenPending = false;    // 'listen' sent, 'listening' event not yet back
let capturing = false;        // meeting capture running
let meetingStarting = false;  // 'meeting' sent, 'meeting-started' not yet back
let meetingTranscript = null;
let meetingAutoStarted = false;
let meetingStopping = false;
let meetingSuppressed = false;
let meetingFreeChecks = 0;
let suppressFreeChecks = 0;
let micCheckInFlight = false;
let lastMicInUse = false;
let agentBusy = false;
let pendingJobs = [];
let agentCooldownUntil = 0;   // set when the Claude usage limit is hit
let cooldownTimer = null;
let ttsProc = null;
let sessionId = null;
let sessionLast = 0;
let lastBriefDate = null;
let quitting = false;
let state = 'idle';
let tray = null;

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
    win.webContents.setZoomFactor(scale);
    send('momo:profile', { character: CHAR, name: CHAR_NAME });
  });
  win.once('ready-to-show', () => {
    setTimeout(() => speak("Hi, I'm " + CHAR_NAME + "! Tap the mic whenever you need me."), 800);
  });
}

// ---------- show / hide (global hotkey + tray; hiding never quits) ----------
function ensureVisible() {
  if (win && !win.isDestroyed() && !win.isVisible()) {
    win.show();
    win.setAlwaysOnTop(true, 'screen-saver');
  }
}
function hideMomo() {
  if (!win || win.isDestroyed()) return;
  stopSpeaking();
  cancelListening();
  if (state !== 'meeting') setState(capturing ? 'meeting' : 'idle');
  win.hide();          // background work (meetings, reminders, brief) keeps running
}
function showMomo(listen) {
  if (!win || win.isDestroyed()) return;
  ensureVisible();
  if (listen && !capturing) startListening();   // summoned = ready to hear you
}
function toggleMomo() {
  if (win && win.isVisible()) hideMomo();
  else showMomo(true);
}

// ---------- TTS: Microsoft neural voice (edge-tts, warm & bilingual) with offline SAPI fallback ----------
function speak(text, onDone) {
  const clean = String(text)
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/[*_#`>|]/g, ' ')
    .replace(/\[[^\]]*\]\([^)]*\)/g, ' ')
    .replace(/https?:\S+/g, ' a link ')
    .replace(/\s+/g, ' ').trim();
  if (!clean) { if (onDone) onDone(); return; }
  ensureVisible();   // she may be hidden when a reminder or reply arrives
  // Stay silent out loud when her voice could reach a call or her own open mic.
  if (capturing || meetingStarting || lastMicInUse || listening || listenPending) {
    bubble(clean, true);
    if (capturing) setState('meeting');
    else if (state === 'thinking' || state === 'speaking') setState('idle');
    if (onDone) onDone();
    return;
  }
  stopSpeaking();
  setState('speaking');
  bubble(clean, true);
  if ((config.tts.engine || 'edge') === 'edge') speakEdge(clean, onDone);
  else speakSapi(clean, onDone);
}
function ttsFinish(p, onDone) {
  return () => {
    if (ttsProc !== p) return;              // a newer utterance took over
    ttsProc = null;
    if (state === 'speaking') setState(capturing ? 'meeting' : 'idle');
    if (onDone) onDone();
  };
}
function speakEdge(clean, onDone) {
  const mp3 = path.join(os.tmpdir(), 'momo-tts-' + Date.now() + '.mp3');
  const b64 = Buffer.from(clean, 'utf8').toString('base64');
  // Route by language: a Hindi voice reads digits/dates in Hindi even inside
  // English sentences ("14" -> chaudah), which sounds wrong. English replies get
  // the Indian-English voice; only Devanagari replies get the Hindi voice.
  const isHindi = /[ऀ-ॿ]/.test(clean);
  // per-character defaults: Momo = female voices, Toto = male voices
  const defEn = CHAR === 'toto' ? 'en-IN-PrabhatNeural' : 'en-IN-NeerjaNeural';
  const defHi = CHAR === 'toto' ? 'hi-IN-MadhurNeural' : 'hi-IN-SwaraNeural';
  const voice = config.tts.edgeVoice ||
    (isHindi ? (config.tts.edgeVoiceHindi || defHi)
             : (config.tts.edgeVoiceEnglish || defEn));
  const p1 = spawn(PYTHON_EXE, [path.join(PROJECT, 'python', 'momo_voice.py'),
    '--voice', voice,
    '--rate', config.tts.edgeRate || '+0%',
    '--out', mp3, '--text-b64', b64],
    { windowsHide: true, env: Object.assign({}, process.env, { PYTHONUTF8: '1' }) });
  ttsProc = p1;
  const t1 = setTimeout(() => { try { p1.kill(); } catch (e) {} }, 25000);
  p1.on('exit', code => {
    clearTimeout(t1);
    if (ttsProc !== p1) { try { fs.unlinkSync(mp3); } catch (e) {} return; }   // cancelled
    if (code !== 0 || !fs.existsSync(mp3)) {
      log('edge-tts failed, falling back to offline voice');
      speakSapi(clean, onDone);            // reassigns ttsProc itself
      return;
    }
    const ps = [
      'Add-Type -AssemblyName PresentationCore;',
      '$p = New-Object System.Windows.Media.MediaPlayer;',
      "$p.Open([uri]('file:///' + '" + mp3.replace(/\\/g, '/') + "'));",
      '$p.Play(); $t = 0;',
      'while (-not $p.NaturalDuration.HasTimeSpan -and $t -lt 100) { Start-Sleep -Milliseconds 100; $t++ };',
      'if ($p.NaturalDuration.HasTimeSpan) { Start-Sleep -Milliseconds ([int]$p.NaturalDuration.TimeSpan.TotalMilliseconds + 300) };',
      '$p.Close()'
    ].join(' ');
    const p2 = spawn('powershell.exe', ['-NoProfile', '-Sta', '-ExecutionPolicy', 'Bypass', '-Command', ps], { windowsHide: true });
    ttsProc = p2;
    const done = ttsFinish(p2, onDone);
    p2.on('exit', () => { try { fs.unlinkSync(mp3); } catch (e) {} done(); });
    p2.on('error', e => { log('tts play error', e.message); done(); });
  });
  p1.on('error', e => {
    clearTimeout(t1);
    log('edge-tts spawn error', e.message);
    if (ttsProc === p1) speakSapi(clean, onDone);
  });
}
function speakSapi(clean, onDone) {
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
  const done = ttsFinish(p, onDone);
  p.on('exit', done);
  p.on('error', e => { log('tts error', e.message); done(); });
}
function stopSpeaking() {
  const p = ttsProc;
  ttsProc = null;
  if (p) { try { p.kill(); } catch (e) {} }
}

// ---------- Ear daemon ----------
// Resolve the REAL python interpreter once: 'python' on Windows is often a
// launcher shim whose extra process breaks stdin control and orphans the
// daemon when killed.
let PYTHON_EXE = 'python';
try {
  const real = execSync('python -c "import sys; print(sys.executable)"',
    { windowsHide: true, timeout: 15000 }).toString().trim();
  if (real && fs.existsSync(real)) PYTHON_EXE = real;
} catch (e) { /* fall back to 'python' */ }

function earSend(cmd) {
  if (!ear || !ear.stdin.writable) return false;
  try { ear.stdin.write(cmd + '\n'); return true; } catch (e) { return false; }
}
function startEarDaemon() {
  if (ear) return;
  const args = [path.join(PROJECT, 'python', 'momo_ear.py'),
    '--model', config.stt.modelDir,
    '--silence-ms', String(config.stt.silenceMs || 1800),
    '--engine', config.stt.commandEngine || 'whisper',
    '--whisper-model', config.stt.whisperModel || 'small',
    '--chunk-sec', String(config.stt.chunkSec || 30),
    '--glossary-file', path.join(DATA, 'meeting_glossary.txt')];
  if (config.stt.inputDevice) args.push('--input-device', config.stt.inputDevice);
  const p = spawn(PYTHON_EXE, args, {
    cwd: PROJECT, windowsHide: true,
    env: Object.assign({}, process.env, { PYTHONUTF8: '1' })
  });
  ear = p;
  earReady = false;
  let buf = '';
  p.stdout.on('data', d => {
    if (ear !== p) return;
    buf += d.toString();
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line) continue;
      let msg; try { msg = JSON.parse(line); } catch (e) { continue; }
      onEarEvent(msg);
    }
  });
  p.stderr.on('data', d => log('ear:', d.toString().trim()));
  const earDied = (why) => {
    if (ear !== p) return;
    ear = null; earReady = false; listening = false; listenPending = false;
    meetingStarting = false;
    log('ear daemon gone:', why);
    if (quitting) return;
    if (capturing) {                        // crashed mid-meeting: salvage the notes
      capturing = false;
      send('momo:rec', false);
      bubble('Note-taking stopped unexpectedly — writing up what I have.', true);
      finishMeetingNotes();
    } else if (state === 'listening') {
      setState('idle');
    }
    const delay = Math.min(30000, 2000 * Math.pow(2, Math.min(earRestarts++, 4)));
    setTimeout(startEarDaemon, delay);      // keep her ears alive
  };
  p.on('exit', code => earDied('exit ' + code));
  p.on('error', e => {
    log('ear spawn error', e.message);
    bubble('Mic engine failed to start: ' + e.message, true);
    // 'exit' usually follows for a started process; if it never comes (spawn
    // failure), clean up ourselves after a beat
    setTimeout(() => earDied('error ' + e.message), 2000);
  });
}
function onEarEvent(msg) {
  switch (msg.event) {
    case 'daemon-ready':
      earReady = true; earRestarts = 0;
      log('ear ready, device:', msg.device);
      break;
    case 'listening':
      listening = true;
      listenPending = false;
      setState('listening');
      bubble('Listening…');
      break;
    case 'partial':
      send('momo:partial', msg.text);
      break;
    case 'transcribing':
      send('momo:partial', '');
      bubble('Getting that…', true);
      break;
    case 'final':
      if (!listening) break;       // this listen cycle was cancelled: don't act on it
      listening = false;
      listenPending = false;
      if (msg.text) handleCommand(msg.text);
      break;
    case 'cancelled':
      listening = false;
      listenPending = false;
      if (state === 'listening') setState(capturing ? 'meeting' : 'idle');
      break;
    case 'meeting-started':
      meetingStarting = false;
      capturing = true;
      meetingTranscript = msg.transcript;
      send('momo:rec', true);
      setState('meeting');
      if (meetingSuppressed) stopMeeting();   // user declined while it was starting
      break;
    case 'meeting-stopped':
      capturing = false;
      meetingStopping = false;
      meetingAutoStarted = false;
      send('momo:rec', false);
      finishMeetingNotes();
      break;
    case 'status':
      log('ear status:', msg.message);
      break;
    case 'error':
      log('ear error:', msg.message);
      bubble('Mic problem: ' + msg.message, true);
      if (meetingStarting) {
        // the capture never came up (no mic etc.) — don't let the watcher
        // retry every tick for the rest of the call
        meetingStarting = false;
        meetingAutoStarted = false;
        meetingSuppressed = true;
        suppressFreeChecks = 0;
      }
      if (listenPending) { listenPending = false; if (state === 'listening') setState('idle'); }
      break;
  }
}
function earCommand(cmd) {
  if (earSend(cmd)) return true;
  earReady = false;
  bubble('My ears just hiccuped — restarting them…', true);
  startEarDaemon();
  return false;
}
function startListening() {
  if (listening || listenPending) return;
  if (!earReady) { bubble('My ears are still waking up — one second…', true); startEarDaemon(); return; }
  stopSpeaking();
  listenPending = true;
  if (!earCommand('listen')) listenPending = false;
}
function cancelListening() {
  if (!listening && !listenPending) return;
  listening = false;
  listenPending = false;
  earCommand('cancel');
}

// ---------- Meeting capture ----------
function startMeeting(auto) {
  if (capturing || meetingStarting || !earReady) { if (!earReady) startEarDaemon(); return; }
  stopSpeaking();               // never let an in-flight utterance play into the call
  cancelListening();
  const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
  const file = path.join(DATA, 'transcripts', 'meeting-' + stamp + '.txt');
  meetingAutoStarted = !!auto;
  meetingStarting = true;
  if (!earCommand('meeting ' + file)) { meetingStarting = false; return; }
  ensureVisible();   // the REC badge must be seen while notes are being taken
  if (auto) bubble('Meeting detected — taking notes quietly.', true);
  else speak('Taking meeting notes from now.');
}
function stopMeeting() {
  if (meetingStopping) return;
  if (!capturing) {
    // a start may still be in flight; meeting-started will see meetingSuppressed
    // and stop it — nothing to send yet
    return;
  }
  meetingStopping = true;
  setState('thinking');
  bubble('Meeting ended — finishing the transcript…', true);
  earCommand('meeting-stop');
}
function finishMeetingNotes() {
  const t = meetingTranscript;
  meetingTranscript = null;
  meetingStopping = false;
  const empty = !t || !fs.existsSync(t) || fs.statSync(t).size < 200;
  if (empty) {
    if (!listening) setState('idle');
    bubble('Meeting ended — nothing much was said, so no notes this time.');
    return;
  }
  setState('thinking');
  bubble('Writing up your meeting notes…', true);
  runAgent('[MEETING-NOTES] The meeting just ended. Transcript file: ' + t + '. Follow knowledge/meeting-notes.md.',
    { fresh: true, queue: true }, reply => speak(reply));
}

// ---------- Auto meeting detection ----------
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
  let out = '', settled = false;
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
      if (!capturing && !meetingStarting && !meetingStopping && !meetingSuppressed) startMeeting(true);
    } else {
      if (meetingSuppressed) {
        suppressFreeChecks++;
        if (suppressFreeChecks >= (config.meeting.endAfterChecks || 2)) {
          meetingSuppressed = false; suppressFreeChecks = 0;
        }
      }
      if (capturing && meetingAutoStarted && !meetingStopping) {
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
function scheduleCooldownDrain() {
  if (cooldownTimer) return;
  cooldownTimer = setTimeout(() => {
    cooldownTimer = null;
    drainJobs();
  }, Math.max(1000, agentCooldownUntil - Date.now() + 1000));
}
function runAgent(prompt, opts, cb) {
  opts = opts || {};
  if (Date.now() < agentCooldownUntil) {
    // usage limit hit earlier: don't burn attempts; queue important work for later
    if (opts.queue) { pendingJobs.push({ prompt, opts, cb }); scheduleCooldownDrain(); return; }
    if (!opts.background) {
      const mins = Math.ceil((agentCooldownUntil - Date.now()) / 60000);
      speak('My thinking limit is reached for now — I should be back in about ' + mins + ' minutes.');
    }
    return;
  }
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
  const model = opts.model || config.agent.model;   // per-task routing (heartbeats run cheap)
  if (model) args.push('--model', model);
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
      reply = j.result ? String(j.result).trim() : null;
      if (!opts.background) { sessionId = j.session_id || sessionId; sessionLast = Date.now(); }
    } catch (e) {
      log('agent parse fail. code=', code, 'err=', err.slice(0, 500), 'out=', out.slice(0, 500));
    }
    if (!reply) {
      // usage limit? degrade gracefully instead of dying silently
      const blob = (out + ' ' + err).toLowerCase();
      if (/usage limit|rate.?limit|overloaded|credit balance|insufficient credit/.test(blob)) {
        agentCooldownUntil = Date.now() + 30 * 60000;   // hold off for 30 min
        scheduleCooldownDrain();
        log('usage limit hit — cooling down until', new Date(agentCooldownUntil).toString());
        if (!opts.background) speak('I have used up my thinking limit for now. I will finish pending work as soon as it resets — around half an hour.');
        else bubble('Thinking limit reached — background checks paused for a while.', true);
        return;
      }
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
  runAgent(text, { queue: true }, reply => {
    speak(reply, () => {
      // If Momo asked a question, reopen the mic so the user can just answer —
      // but never during a live call (the call's audio would become a command).
      if (/\?\s*$/.test(reply) && !capturing && !meetingStarting && !lastMicInUse) startListening();
    });
  });
}

// ---------- Schedulers ----------
function heartbeatTick() {
  if (agentBusy || capturing || listening || state === 'speaking') return;
  if (!config.heartbeatMinutes) return;
  runAgent('[HEARTBEAT] Time: ' + new Date().toString() + '. Follow the heartbeat procedure in knowledge/reminders-todos.md. Reply SILENT if nothing is due.',
    { fresh: true, background: true, model: config.agent.heartbeatModel || 'haiku' }, reply => {
      if (reply && reply.toUpperCase() !== 'SILENT') speak(reply);
    });
}
function briefTick() {
  const now = new Date();
  const hhmm = now.toTimeString().slice(0, 5);
  const today = now.toDateString();
  if (hhmm === config.briefTime && lastBriefDate !== today &&
      !agentBusy && !capturing && !listening && state !== 'speaking') {
    lastBriefDate = today;
    setState('thinking');
    bubble('Preparing your morning brief…', true);
    runAgent('[BRIEF] Morning brief for ' + now.toDateString() + '. Follow knowledge/daily-brief.md.',
      { fresh: true }, reply => speak(reply));
  }
}

// ---------- IPC ----------
ipcMain.on('mic-toggle', () => {
  if (ttsProc) { stopSpeaking(); setState(capturing ? 'meeting' : 'idle'); return; }
  if (listening || listenPending) { cancelListening(); bubble('Okay, cancelled.'); }
  else startListening();
});
ipcMain.on('meeting-toggle', () => {
  if (capturing || meetingStarting) {
    // stop — including a start still in flight (meeting-started will honor
    // meetingSuppressed and stop immediately)
    meetingSuppressed = true;
    suppressFreeChecks = 0;
    stopMeeting();
    if (meetingStarting) bubble('Okay — no notes for this meeting.', true);
  } else {
    startMeeting(false);
  }
});
ipcMain.on('text-command', (e, text) => { if (text && text.trim()) handleCommand(text.trim()); });
// the ✖ button QUITS the app entirely; Alt+M is the hide/show toggle
ipcMain.on('momo-quit', () => { app.quit(); });

// ---------- app lifecycle ----------
if (!app.requestSingleInstanceLock()) app.quit();
// Ctrl+Alt+M is owned by the Start Menu shortcut: with Momo closed it launches her;
// with Momo running the relaunch arrives here as a second-instance signal -> toggle.
app.on('second-instance', () => toggleMomo());
app.whenReady().then(() => {
  createWindow();
  startEarDaemon();                                   // models load once, up front

  // global hotkey: summon (and listen) / dismiss from anywhere
  const hotkey = config.ui.hotkey || 'Alt+M';
  try {
    if (!globalShortcut.register(hotkey, toggleMomo)) log('hotkey busy:', hotkey);
    else log('hotkey registered:', hotkey);
  } catch (e) { log('hotkey error:', e.message); }

  // tray icon: the only place to really quit
  try {
    const icon = nativeImage.createFromPath(path.join(__dirname, 'renderer', CHAR + '.png'))
      .resize({ width: 24, height: 24 });
    tray = new Tray(icon);
    tray.setToolTip(CHAR_NAME + '  (' + hotkey + ')');
    tray.setContextMenu(Menu.buildFromTemplate([
      { label: 'Show / hide  (' + hotkey + ')', click: toggleMomo },
      { type: 'separator' },
      { label: 'Quit ' + CHAR_NAME, click: () => app.quit() },
    ]));
    tray.on('click', toggleMomo);
  } catch (e) { log('tray error:', e.message); }

  setInterval(heartbeatTick, Math.max(5, config.heartbeatMinutes || 15) * 60000);
  setInterval(briefTick, 30000);
  setInterval(meetingWatchTick, Math.max(5, config.meeting.checkSeconds || 15) * 1000);
});
app.on('will-quit', () => { try { globalShortcut.unregisterAll(); } catch (e) {} });
app.on('before-quit', (e) => {
  if (quitting) return;               // second pass: let the quit proceed
  quitting = true;
  stopSpeaking();
  if (ear) {
    // hold the quit briefly so the daemon can flush its transcript
    e.preventDefault();
    const p = ear;
    earSend('quit');
    const finish = () => { try { p.kill(); } catch (e2) {} app.exit(0); };
    p.once('exit', () => app.exit(0));
    setTimeout(finish, 1500);
  }
});
app.on('window-all-closed', () => app.quit());
