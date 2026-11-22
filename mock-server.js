/* ?mock=1 — an in-memory stand-in for Drive v3 that speaks the same HTTP shapes the real client sends.
   The app code path is identical to production; only `fetch` is replaced.
   Every request is logged in window.__mockLog; window.__mock exposes test hooks. */
import { FOLDER_MIME, parseRangeHeader } from './lib.js';

const J = (obj, status = 200, headers = {}) =>
  new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json', ...headers } });
const err = (status, message) => J({ error: { code: status, message } }, status);
const iso = () => new Date().toISOString();

export function createMockServer({ speedMs = 90 } = {}) {
  const files = new Map();
  const perms = new Map();   // fileId -> [perm]
  const sessions = new Map(); // uri -> { meta, size, received, chunksSeen }
  let nextId = 1, nextPerm = 1, nextSess = 1;
  const log = [];
  const hooks = { failNext: 0, failStatus: 503, chunkDelay: speedMs, lastEmails: [], offline: false };
  let changeSeq = 0; const changeLog = [];
  const touch = (id) => { changeSeq++; changeLog.push({ seq: changeSeq, fileId: id }); };
  const user = { displayName: 'Mock User', emailAddress: 'mock.user@example.test', photoLink: null };

  files.set('root', { id: 'root', name: 'My Drive', mimeType: FOLDER_MIME, parents: [], trashed: false });

  const mk = (name, mimeType, parents, extra = {}) => {
    const id = 'mock' + (nextId++);
    const f = { id, name, mimeType, parents, trashed: false, createdTime: iso(), modifiedTime: iso(), ...extra };
    f.webViewLink = (mimeType === FOLDER_MIME ? 'https://drive.google.com/drive/folders/' : 'https://drive.google.com/file/d/') + id;
    f.rev = 1; files.set(id, f); touch(id); return f;
  };
  const pub = (f) => { const { trashed, content, rev, ...rest } = f; return { ...rest, headRevisionId: 'rev' + rev, size: f.size == null ? undefined : String(f.size) }; };

  function matchQ(f, q) {
    if (!q) return true;
    const parts = q.split(/ and /i);
    for (const p of parts) {
      let m;
      if ((m = /^'((?:\\.|[^'\\])*)' in parents$/.exec(p.trim()))) { if (!f.parents.includes(m[1].replace(/\\(.)/g, '$1'))) return false; }
      else if ((m = /^name = '((?:\\.|[^'\\])*)'$/.exec(p.trim()))) { if (f.name !== m[1].replace(/\\(.)/g, '$1')) return false; }
      else if ((m = /^mimeType = '([^']*)'$/.exec(p.trim()))) { if (f.mimeType !== m[1]) return false; }
      else if (/^trashed = false$/.test(p.trim())) { if (f.trashed) return false; }
    }
    return true;
  }

  async function handler(url, init = {}) {
    const u = new URL(url, 'https://mock.test/');
    const method = (init.method || 'GET').toUpperCase();
    const headers = init.headers || {};
    let bodyText = typeof init.body === 'string' ? init.body : null;
    const entry = { t: Date.now(), method, url: u.origin + u.pathname + u.search, body: bodyText ? safeParse(bodyText) : (init.body ? '[binary ' + (init.body.size ?? '?') + ' B]' : undefined), range: headers['Content-Range'] };
    log.push(entry);
    if (log.length > 5000) log.shift();
    if (hooks.offline) throw new TypeError('Failed to fetch (mock offline)');
    if (hooks.force401) { hooks.force401 = false; return err(401, 'expired token'); }
    if (init.signal && init.signal.aborted) throw new DOMException('Aborted', 'AbortError');
    if (!/^Bearer /.test(headers.Authorization || '')) return err(401, 'no token');

    let m;
    const path = u.pathname;
    const q = u.searchParams;

    // ---- upload sessions ----
    if (path === '/upload/drive/v3/files' && method === 'POST' && q.get('uploadType') === 'multipart') {
      const B = /boundary=(.+)$/.exec(headers['Content-Type'] || '')[1];
      const parts = bodyText.split('--' + B);
      const CRLF = String.fromCharCode(13, 10), SEP = CRLF + CRLF;
      const trim = (x) => (x.endsWith(CRLF) ? x.slice(0, -2) : x);
      const meta = JSON.parse(trim(parts[1].split(SEP)[1]));
      const content = trim(parts[2].split(SEP).slice(1).join(SEP));
      const f = mk(meta.name, meta.mimeType, meta.parents || ['root'], { size: content.length, content });
      return J(pub(f));
    }
    m = /^\/upload\/drive\/v3\/files\/([^/]+)$/.exec(path);
    if (m && method === 'PATCH') {
      const f = files.get(m[1]); if (!f) return err(404, 'File not found');
      f.content = bodyText; f.size = bodyText.length; f.rev++; f.modifiedTime = iso(); touch(f.id);
      return J(pub(f));
    }
    if (path === '/upload/drive/v3/files' && method === 'POST') {
      const meta = JSON.parse(bodyText || '{}');
      const size = Number(headers['X-Upload-Content-Length']);
      const uri = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&upload_id=mocksess' + (nextSess++);
      sessions.set(uri, { meta, size, received: 0, mime: headers['X-Upload-Content-Type'] });
      return new Response(null, { status: 200, headers: { Location: uri } });
    }
    if (path === '/upload/drive/v3/files' && method === 'PUT') {
      const key = u.origin + u.pathname + '?uploadType=resumable&upload_id=' + q.get('upload_id');
      const s = sessions.get(key);
      if (!s) return err(404, 'session not found');
      const cr = headers['Content-Range'] || '';
      if (/^bytes \*\//.test(cr)) return statusReply(s);
      if (hooks.failNext > 0) { hooks.failNext--; return err(hooks.failStatus, 'injected failure'); }
      const m = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(cr);
      if (!m) return err(400, 'bad Content-Range');
      const start = Number(m[1]), end = Number(m[2]) + 1;
      if (start !== s.received) return err(400, 'offset mismatch: expected ' + s.received + ' got ' + start);
      await new Promise((res, rej) => {
        const t = setTimeout(res, hooks.chunkDelay);
        if (init.signal) init.signal.addEventListener('abort', () => { clearTimeout(t); rej(new DOMException('Aborted', 'AbortError')); }, { once: true });
      });
      s.received = end;
      if (s.received >= s.size) {
        const f = mk(s.meta.name, s.meta.mimeType, s.meta.parents || ['root'], { size: s.size });
        sessions.delete(key);
        return J(pub(f), 200);
      }
      return statusReply(s);
    }
    function statusReply(s) {
      if (s.received >= s.size && s.size === 0) { const f = mk(s.meta.name, s.meta.mimeType, s.meta.parents || ['root'], { size: 0 }); return J(pub(f)); }
      const h = {}; if (s.received > 0) h.Range = 'bytes=0-' + (s.received - 1);
      return new Response(null, { status: 308, headers: h });
    }

    // ---- Drive v3 ----
    if (path === '/drive/v3/changes/startPageToken') return J({ startPageToken: String(changeSeq) });
    if (path === '/drive/v3/changes') {
      const from = Number(q.get('pageToken'));
      return J({ changes: changeLog.filter(c => c.seq > from).map(c => ({ fileId: c.fileId })), newStartPageToken: String(changeSeq) });
    }
    if (path === '/drive/v3/about') return J({ user, storageQuota: { limit: String(15 * 1024 ** 3), usage: String(totalBytes()) } });
    if (path === '/drive/v3/files' && method === 'GET') {
      const list = [...files.values()].filter(f => f.id !== 'root' && matchQ(f, q.get('q')));
      list.sort((a, b) => a.name.localeCompare(b.name));
      return J({ files: list.map(pub) });
    }
    if (path === '/drive/v3/files' && method === 'POST') {
      const b = JSON.parse(bodyText || '{}');
      return J(pub(mk(b.name, b.mimeType || 'application/octet-stream', b.parents || ['root'])));
    }
    m = /^\/drive\/v3\/files\/([^/]+)$/.exec(path);
    if (m) {
      const f = files.get(decodeURIComponent(m[1]));
      if (!f) return err(404, 'File not found');
      if (method === 'GET' && q.get('alt') === 'media') return new Response(f.content != null ? f.content : new Uint8Array(Math.min(f.size || 0, 1e6)), { status: 200 });
      if (method === 'GET') return J(pub(f));
      if (method === 'PATCH') {
        const b = JSON.parse(bodyText || '{}');
        if (b.name != null) f.name = b.name;
        if (b.trashed != null) { f.trashed = !!b.trashed; if (f.trashed) trashChildren(f.id); }
        f.modifiedTime = iso(); f.rev++; touch(f.id);
        return J(pub(f));
      }
    }
    m = /^\/drive\/v3\/files\/([^/]+)\/permissions(?:\/([^/]+))?$/.exec(path);
    if (m) {
      const fid = decodeURIComponent(m[1]);
      if (!files.has(fid)) return err(404, 'File not found');
      const list = perms.get(fid) || (perms.set(fid, []), perms.get(fid));
      if (method === 'GET') return J({ permissions: list });
      if (method === 'POST') {
        const b = JSON.parse(bodyText || '{}');
        if (b.type === 'user' && /rejected@/.test(b.emailAddress || '')) return err(400, 'Invalid recipient (mock)');
        const p = { id: 'perm' + (nextPerm++), type: b.type, role: b.role };
        if (b.type === 'user') { p.emailAddress = b.emailAddress; p.displayName = b.emailAddress.split('@')[0]; hooks.lastEmails.push({ to: b.emailAddress, notify: q.get('sendNotificationEmail'), message: q.get('emailMessage') }); }
        list.push(p); return J(p);
      }
      if (method === 'DELETE') {
        const i = list.findIndex(p => p.id === decodeURIComponent(m[2]));
        if (i < 0) return err(404, 'Permission not found');
        list.splice(i, 1); return new Response(null, { status: 204 });
      }
    }
    return err(404, 'mock: no route ' + method + ' ' + path);
  }

  function trashChildren(id) { for (const f of files.values()) if (f.parents.includes(id)) { f.trashed = true; trashChildren(f.id); } }
  function totalBytes() { let n = 0; for (const f of files.values()) if (!f.trashed && f.size) n += Number(f.size); return n; }
  function safeParse(s) { try { return JSON.parse(s); } catch (e) { return s; } }

  /** Simulates another device: create or overwrite a text file under parentId. */
  function remoteWrite(parentId, name, content) {
    const ex = [...files.values()].find(f => !f.trashed && f.name === name && f.parents.includes(parentId));
    if (ex) { ex.content = content; ex.size = content.length; ex.rev++; ex.modifiedTime = iso(); touch(ex.id); return ex.id; }
    return mk(name, 'application/json', [parentId], { size: content.length, content }).id;
  }
  function remoteDelete(id) { const f = files.get(id); if (f) { f.trashed = true; touch(id); } }
  return { handler, log, hooks, files, perms, sessions, user, remoteWrite, remoteDelete };
}
