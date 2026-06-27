/* Shared state management - BPS Kabupaten Toli-Toli */
const BPS_KEY = 'bps_tolitoli_antrian';
const BPS_CH  = 'bps_antrian_ch';

const SERVICE_MAP = {
  K: { code: 'K', name: 'Konsultasi Statistik',  icon: '💬', desk: 'Meja Konsultasi' },
  P: { code: 'P', name: 'Perpustakaan Statistik', icon: '📚', desk: 'Meja Perpustakaan' },
  R: { code: 'R', name: 'Rekomendasi Statistik',  icon: '📝', desk: 'Meja Rekomendasi' },
  L: { code: 'L', name: 'Pelayanan Lainnya',       icon: '📋', desk: 'Meja Pelayanan' },
};

function _defaultState() {
  return {
    queues:         [],
    counters:       { K: 0, P: 0, R: 0, L: 0 },
    currentServing: null,
    officer:        { name: '', email: '' },
    lastUpdated:    Date.now(),
  };
}

function getState() {
  try {
    const raw = localStorage.getItem(BPS_KEY);
    return raw ? JSON.parse(raw) : _defaultState();
  } catch { return _defaultState(); }
}

function saveState(state) {
  state.lastUpdated = Date.now();
  localStorage.setItem(BPS_KEY, JSON.stringify(state));
  try {
    const ch = new BroadcastChannel(BPS_CH);
    ch.postMessage({ type: 'UPDATE', state });
    ch.close();
  } catch { /* browser may not support */ }
}

function addQueue({ code, name, phone, institution, gender, purpose }) {
  const state = getState();
  state.counters[code] = (state.counters[code] || 0) + 1;
  const num = String(state.counters[code]).padStart(3, '0');
  const svc = SERVICE_MAP[code];
  const waitingBefore = state.queues.filter(q => q.code === code && q.status === 'waiting').length;
  const queue = {
    id: `${code}-${num}`,
    code,
    number: state.counters[code],
    typeName: svc.name,
    desk: svc.desk,
    name, phone, institution, gender, purpose,
    status: 'waiting',
    createdAt: Date.now(),
    calledAt: null,
    doneAt: null,
    waitingBefore,
  };
  state.queues.push(queue);
  saveState(state);
  return queue;
}

function callQueue(id) {
  const state = getState();
  const q = state.queues.find(x => x.id === id);
  if (!q) return null;
  q.status = 'serving';
  q.calledAt = Date.now();
  state.currentServing = { ...q };
  saveState(state);
  return q;
}

function doneQueue(id) {
  const state = getState();
  const q = state.queues.find(x => x.id === id);
  if (!q) return;
  q.status = 'done';
  q.doneAt = Date.now();
  if (state.currentServing && state.currentServing.id === id) {
    state.currentServing = null;
  }
  saveState(state);
}

function skipQueue(id) {
  const state = getState();
  const q = state.queues.find(x => x.id === id);
  if (!q) return;
  q.status = 'skip';
  if (state.currentServing && state.currentServing.id === id) {
    state.currentServing = null;
  }
  saveState(state);
}

function setOfficer(name, email) {
  const state = getState();
  state.officer = { name, email };
  saveState(state);
}

function callNextByCode(code) {
  const state = getState();
  const next = state.queues.find(q => q.code === code && q.status === 'waiting');
  if (!next) return null;
  return callQueue(next.id);
}

function resetAll() {
  saveState(_defaultState());
}

function onStateChange(cb) {
  window.addEventListener('storage', e => {
    if (e.key === BPS_KEY) cb(getState());
  });
  try {
    const ch = new BroadcastChannel(BPS_CH);
    ch.onmessage = e => { if (e.data.type === 'UPDATE') cb(e.data.state); };
  } catch { /* polling fallback handled by caller */ }
}

function formatTime(ts) {
  if (!ts) return '-';
  return new Date(ts).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
}

function formatDate(ts) {
  if (!ts) return '-';
  return new Date(ts).toLocaleDateString('id-ID', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });
}
