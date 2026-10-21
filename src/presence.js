/*
 * "Online now" bar: how many people have this site open, and how many testers came today.
 *
 * Uses Firebase Realtime Database (src/firebase-config.js). Privacy: each browser gets an
 * anonymous random ID (no name, no account); only that ID, a timestamp and the page type
 * ("home" / "view") are stored - never document names or content. Closed tabs are removed
 * by Firebase automatically (onDisconnect).
 *
 * Framework-free so the Vite app and the Angular app can both use it:
 *   startPresence(container, { env: 'staging' })   // or 'angular'
 */
import { FIREBASE_CONFIG } from './firebase-config.js';

const ID_KEY = 'dv-tester-id';

function testerId() {
  try {
    let id = localStorage.getItem(ID_KEY);
    if (!id) { id = crypto.randomUUID().replace(/-/g, '').slice(0, 20); localStorage.setItem(ID_KEY, id); }
    return id;
  } catch {
    return crypto.randomUUID().replace(/-/g, '').slice(0, 20); // private mode: per page load
  }
}

const today = () => new Date().toISOString().slice(0, 10); // UTC day

export function startPresence(container, { env }) {
  if (!FIREBASE_CONFIG?.databaseURL || !container) return { setPage() {} };

  const bar = document.createElement('div');
  bar.className = 'dv-presence';
  bar.setAttribute('role', 'status');
  bar.hidden = true;
  container.append(bar);

  let page = 'home'; // 'home' or 'view'; kept until Firebase is connected
  let setPage = (p) => { page = p; };
  const api = { setPage: (p) => setPage(p) };

  // Load Firebase after the page has finished loading, so the viewer itself stays fast.
  const start = async () => {
    try {
      const [{ initializeApp }, db] = await Promise.all([import('firebase/app'), import('firebase/database')]);
      const app = initializeApp(FIREBASE_CONFIG, 'docs-viewer-presence');
      const database = db.getDatabase(app);
      const id = testerId();
      const tab = Math.random().toString(36).slice(2, 10);
      const me = db.ref(database, `presence/${env}/${id}_${tab}`);
      let online = 0;
      let todayCount = 0;
      const render = () => {
        bar.hidden = false;
        bar.innerHTML = '';
        const dot = document.createElement('span'); dot.className = 'dv-presence-dot';
        const text = document.createElement('span');
        text.textContent = `${online} online now · ${todayCount} tester${todayCount === 1 ? '' : 's'} today`;
        bar.append(dot, text);
        bar.title = `People with this ${env} site open right now, and different browsers that opened it today (UTC).`;
      };

      // my presence: written when connected, removed by the server when the tab goes away
      db.onValue(db.ref(database, '.info/connected'), async (snap) => {
        if (snap.val() !== true) return;
        await db.onDisconnect(me).remove();
        await db.set(me, { t: db.serverTimestamp(), page });
      });
      setPage = (p) => { page = p; db.update(me, { page: p }).catch(() => {}); };

      // today's testers: one entry per browser per day (written once; the rules refuse rewrites)
      const day = today();
      const visit = db.ref(database, `visits/${env}/${day}/${id}`);
      db.get(visit).then((s) => (s.exists() ? null : db.set(visit, true))).catch(() => {});

      db.onValue(db.ref(database, `presence/${env}`), (snap) => { online = snap.size; render(); });
      db.onValue(db.ref(database, `visits/${env}/${day}`), (snap) => { todayCount = snap.size; render(); });
    } catch (err) {
      console.warn('Online counter unavailable.', err);
      bar.remove();
    }
  };
  // Start when the browser is idle after the page is ready (documents keep loading meanwhile).
  const later = () => (window.requestIdleCallback ? requestIdleCallback(start, { timeout: 2000 }) : setTimeout(start, 300));
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', later, { once: true });
  else later();
  return api;
}
