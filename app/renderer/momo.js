// Momo renderer — character states, bubble, controls
const bubbleEl = document.getElementById('bubble');
const bubbleText = document.getElementById('bubble-text');
const typebox = document.getElementById('typebox');
const typeinput = document.getElementById('typeinput');
let hideTimer = null;

function showBubble(html, sticky) {
  bubbleText.innerHTML = html;
  bubbleEl.classList.remove('hidden');
  if (hideTimer) clearTimeout(hideTimer);
  if (!sticky) hideTimer = setTimeout(() => bubbleEl.classList.add('hidden'), 6000);
}

window.momo.onState((s) => {
  document.body.classList.remove('listening', 'thinking', 'speaking', 'meeting');
  if (s && s !== 'idle') document.body.classList.add(s);
  if (s === 'thinking') {
    showBubble('Thinking <span class="dots"><i></i><i></i><i></i></span>', true);
  }
  if (s === 'idle') {
    if (hideTimer) clearTimeout(hideTimer);
    hideTimer = setTimeout(() => bubbleEl.classList.add('hidden'), 5000);
  }
});

window.momo.onBubble((b) => showBubble(escapeHtml(b.text), b.sticky));

// character comes from data/profile.json (momo = pink, toto = blue) —
// body image AND accent theme both follow the profile
window.momo.onProfile((p) => {
  const r = document.getElementById('robot');
  r.src = p.character + '.png';
  r.alt = p.name;
  document.body.classList.remove('char-momo', 'char-toto');
  document.body.classList.add('char-' + p.character);
});

// REC badge tracks actual capture, independent of the animation state —
// it must never lie about whether the meeting is being transcribed.
window.momo.onRec((on) => document.body.classList.toggle('rec', !!on));

window.momo.onPartial((t) => {
  if (t) showBubble('<span class="partial">' + escapeHtml(t) + '…</span>', true);
});

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

document.getElementById('btn-mic').addEventListener('click', () => window.momo.micToggle());
document.getElementById('btn-notes').addEventListener('click', () => window.momo.meetingToggle());
document.getElementById('btn-close').addEventListener('click', () => window.momo.quit());
document.getElementById('btn-kb').addEventListener('click', () => {
  typebox.classList.toggle('hidden');
  if (!typebox.classList.contains('hidden')) typeinput.focus();
});
typeinput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && typeinput.value.trim()) {
    window.momo.sendText(typeinput.value.trim());
    typeinput.value = '';
    typebox.classList.add('hidden');
  }
  if (e.key === 'Escape') typebox.classList.add('hidden');
});
