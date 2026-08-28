const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('momo', {
  micToggle: () => ipcRenderer.send('mic-toggle'),
  meetingToggle: () => ipcRenderer.send('meeting-toggle'),
  sendText: (t) => ipcRenderer.send('text-command', t),
  quit: () => ipcRenderer.send('momo-quit'),
  onState: (fn) => ipcRenderer.on('momo:state', (e, s) => fn(s)),
  onBubble: (fn) => ipcRenderer.on('momo:bubble', (e, b) => fn(b)),
  onPartial: (fn) => ipcRenderer.on('momo:partial', (e, t) => fn(t)),
  onRec: (fn) => ipcRenderer.on('momo:rec', (e, b) => fn(b)),
  onProfile: (fn) => ipcRenderer.on('momo:profile', (e, p) => fn(p))
});
