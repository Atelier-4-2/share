/* ?mock=1 — in-memory stand-in for Microsoft Graph v1.0 (OneDrive). Same idea as mock-server.js:
   the real OneDrive adapter runs unchanged against this fake `fetch`. */
import { isGraphAligned } from './lib.js';

const J = (obj, status = 200, headers = {}) =>
  new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json', ...headers } });
const err = (status, code, message) => J({ error: { code, message } }, status);
const iso = () => new Date().toISOString();

export function createMockGraph({ speedMs = 90 } = {}) {
  const items = new Map();
  const perms = new Map();
  const sessions = new Map();
  let nextId = 1, nextPerm = 1, nextSess = 1;
  const log = [];
  const hooks = { failNext: 0, failStatus: 503, chunkDelay: speedMs, denyAnonymous: false, invites: [], offline: false };
  const user = { displayName: 'Mock Student', mail: 'mock.student@school.example.test' };

  items.set('approot', { id: 'approot', name: 'Atelier 4·2 Share', folder: {}, parent: null, createdDateTime: iso() });
  const mk = (name, parent, extra) => {
    const id = 'od' + (nextId++);
    const it = { id, name, parent, createdDateTime: iso(), lastModifiedDateTime: iso(), webUrl: 'https://onedrive.live.example/' + id, eTag: '"e' + (nextId) + ',1"', ver: 1, ...extra };
    items.set(id, it); return it;
  };
  const pub = (it) => { const { parent, content, ver, ...r } = it; return r; };
  const bump = (it) => { it.ver = (it.ver || 1) + 1; it.eTag = '"' + it.id + ',' + it.ver + '"'; it.lastModifiedDateTime = iso(); };
  const kids = (id) => [...items.values()].filter(i => i.parent === id);
  const totalBytes = () => [...items.values()].reduce((n, i) => n + (i.size || 0), 0);

  async function handler(url, init = {}) {
    const u = new URL(url, 'https://mock.test/');
    const method = (init.method || 'GET').toUpperCase();
    const headers = init.headers || {};
    const bodyText = typeof init.body === 'string' ? init.body : null;
    let parsed; try { parsed = bodyText ? JSON.parse(bodyText) : undefined; } catch (e) { parsed = bodyText; }
    log.push({ t: Date.now(), method, url: u.origin + u.pathname + u.search, body: parsed !== undefined ? parsed : (init.body ? '[binary ' + (init.body.size ?? '?') + ' B]' : undefined), range: headers['Content-Range'], auth: !!headers.Authorization });
    if (log.length > 5000) log.shift();
    if (hooks.offline) throw new TypeError('Failed to fetch (mock offline)');
    if (hooks.force401) { hooks.force401 = false; return err(401, 'InvalidAuthenticationToken', 'expired token'); }
    if (init.signal && init.signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const path = decodeURIComponent(u.pathname);

    // ---- upload session URL (pre-authorised; Authorization header must NOT be sent) ----
    if (u.host === 'upload.graph.example') {
      if (headers.Authorization) return err(401, 'unauthenticated', 'upload url must not carry Authorization');
      const s = sessions.get(u.pathname.slice(1));
      if (!s) return err(404, 'itemNotFound', 'session expired');
      if (method === 'GET') return J({ nextExpectedRanges: (s.size != null && s.received >= s.size) ? [] : [s.received + '-'], expirationDateTime: iso() });
      if (method === 'PUT') {
        if (hooks.failNext > 0) { hooks.failNext--; return err(hooks.failStatus, 'serviceNotAvailable', 'injected failure'); }
        const m = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(headers['Content-Range'] || '');
        if (!m) return err(416, 'invalidRange', 'bad Content-Range');
        if (s.size == null) s.size = Number(m[3]);
        const start = Number(m[1]), end = Number(m[2]) + 1;
        const len = end - start;
        if (start !== s.received) return err(416, 'invalidRange', 'expected ' + s.received);
        if (end < s.size && !isGraphAligned(len)) return err(400, 'invalidRange', 'fragment must be a multiple of 320 KiB');
        await new Promise((res, rej) => {
          const t = setTimeout(res, hooks.chunkDelay);
          if (init.signal) init.signal.addEventListener('abort', () => { clearTimeout(t); rej(new DOMException('Aborted', 'AbortError')); }, { once: true });
        });
        s.received = end;
        if (end >= s.size) {
          const it = mk(s.name, s.parent, { size: s.size, file: { mimeType: 'application/octet-stream' } });
          sessions.delete(u.pathname.slice(1));
          return J(pub(it), 201);
        }
        return J({ nextExpectedRanges: [end + '-'] }, 202);
      }
    }

    if (!/^Bearer /.test(headers.Authorization || '')) return err(401, 'InvalidAuthenticationToken', 'no token');
    const p = path.replace(/^\/v1\.0/, '');

    // ---- small-file content: create (conflictBehavior=fail) and overwrite (If-Match) ----
    if (p.startsWith('/me/drive/items/') && p.endsWith(':/content') && method === 'PUT') {
      const inner = p.slice('/me/drive/items/'.length, -':/content'.length);
      const cut = inner.indexOf(':/');
      const parent = inner.slice(0, cut), name = inner.slice(cut + 2);
      if (!items.has(parent)) return err(404, 'itemNotFound', 'parent');
      const ex = kids(parent).find(i => i.name === name);
      if (ex && (u.searchParams.get('@microsoft.graph.conflictBehavior') === 'fail')) return err(409, 'nameAlreadyExists', 'exists');
      if (ex) { ex.content = bodyText; ex.size = bodyText.length; bump(ex); return J(pub(ex)); }
      return J(pub(mk(name, parent, { size: bodyText.length, content: bodyText, file: { mimeType: 'application/json' } })), 201);
    }
    if (p.startsWith('/me/drive/items/') && p.endsWith('/content') && !p.includes(':/')) {
      const id = p.slice('/me/drive/items/'.length, -'/content'.length);
      const it = items.get(id);
      if (!it) return err(404, 'itemNotFound', 'no item');
      if (method === 'GET') return new Response(it.content != null ? it.content : new Uint8Array(Math.min(it.size || 0, 1e6)), { status: 200 });
      if (method === 'PUT') {
        const im = headers['If-Match'];
        if (im && im !== it.eTag) return err(412, 'preconditionFailed', 'eTag mismatch');
        it.content = bodyText; it.size = bodyText.length; bump(it);
        return J(pub(it));
      }
    }
    // delta is refused on purpose here (as school tenants may do): the client must fall back to eTag comparison
    if (p.endsWith('/delta')) return err(501, 'notSupported', 'delta not supported (mock)');
    if (p === '/me') return J({ displayName: user.displayName, mail: user.mail, userPrincipalName: user.mail });
    if (p === '/me/drive') return J({ quota: { used: totalBytes(), total: 5 * 1024 ** 3 } });
    if (p === '/me/drive/special/approot') return J(pub(items.get('approot')));

    let m = /^\/me\/drive\/items\/([^/:]+):\/(.+):\/createUploadSession$/.exec(p);
    if (m && method === 'POST') {
      if (!items.has(m[1])) return err(404, 'itemNotFound', 'parent');
      const key = 'mocksess' + (nextSess++);
      sessions.set(key, { name: m[2], parent: m[1], size: null, received: 0 });
      return J({ uploadUrl: 'https://upload.graph.example/' + key, expirationDateTime: iso() });
    }
    m = /^\/me\/drive\/items\/([^/]+)\/children$/.exec(p);
    if (m) {
      if (!items.has(m[1])) return err(404, 'itemNotFound', 'no item');
      if (method === 'GET') return J({ value: kids(m[1]).map(pub) });
      if (method === 'POST') {
        if (kids(m[1]).some(i => i.name === parsed.name)) return err(409, 'nameAlreadyExists', 'exists');
        return J(pub(mk(parsed.name, m[1], { folder: {} })), 201);
      }
    }
    m = /^\/me\/drive\/items\/([^/]+)$/.exec(p);
    if (m) {
      const it = items.get(m[1]);
      if (!it) return err(404, 'itemNotFound', 'no item');
      if (method === 'GET') return J(pub(it));
      if (method === 'PATCH') { if (parsed.name) it.name = parsed.name; bump(it); return J(pub(it)); }
      if (method === 'DELETE') { const rm = (id) => { kids(id).forEach(k => rm(k.id)); items.delete(id); }; rm(it.id); return new Response(null, { status: 204 }); }
    }
    m = /^\/me\/drive\/items\/([^/]+)\/(invite|createLink|permissions)(?:\/([^/]+))?$/.exec(p);
    if (m) {
      if (!items.has(m[1])) return err(404, 'itemNotFound', 'no item');
      const list = perms.get(m[1]) || (perms.set(m[1], [{ id: 'owner1', roles: ['owner'], grantedToV2: { user: { displayName: user.displayName, email: user.mail } } }]), perms.get(m[1]));
      if (m[2] === 'permissions' && method === 'GET') return J({ value: list });
      if (m[2] === 'permissions' && method === 'DELETE') {
        const i = list.findIndex(x => x.id === m[3]);
        if (i < 0) return err(404, 'itemNotFound', 'no permission');
        list.splice(i, 1); return new Response(null, { status: 204 });
      }
      if (m[2] === 'invite' && method === 'POST') {
        const out = [];
        for (const r of parsed.recipients || []) {
          const perm = { id: 'op' + (nextPerm++), roles: parsed.roles || ['read'], invitation: { email: r.email, signInRequired: !!parsed.requireSignIn }, grantedToV2: { user: { displayName: r.email.split('@')[0], email: r.email } } };
          list.push(perm); out.push(perm);
          hooks.invites.push({ to: r.email, sendInvitation: parsed.sendInvitation, message: parsed.message });
        }
        return J({ value: out });
      }
      if (m[2] === 'createLink' && method === 'POST') {
        if (parsed.scope === 'anonymous' && hooks.denyAnonymous) return err(403, 'accessDenied', 'Anonymous links are disabled by your organization (mock)');
        const perm = { id: 'op' + (nextPerm++), roles: [parsed.type === 'edit' ? 'write' : 'read'], link: { type: parsed.type, scope: parsed.scope, webUrl: 'https://1drv.example/s/' + parsed.scope + '/' + m[1] } };
        list.push(perm); return J(perm, 201);
      }
    }
    return err(404, 'notFound', 'mock graph: no route ' + method + ' ' + p);
  }
  /** Simulates another device: create or overwrite a text file under parent. */
  function remoteWrite(parent, name, content) {
    const ex = kids(parent).find(i => i.name === name);
    if (ex) { ex.content = content; ex.size = content.length; bump(ex); return ex.id; }
    return mk(name, parent, { size: content.length, content, file: { mimeType: 'application/json' } }).id;
  }
  function remoteDelete(id) { items.delete(id); }
  return { handler, log, hooks, items, perms, sessions, user, remoteWrite, remoteDelete };
}
