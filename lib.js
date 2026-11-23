/* Atelier 4·2 Share — pure helpers (no DOM, no network). Tested with node --test. */

export const CHUNK_SIZE = 8 * 1024 * 1024; // must be a multiple of 256 KiB
export const FOLDER_MIME = 'application/vnd.google-apps.folder';
export const MAX_RETRIES = 6;

/** Number of chunks for a file of `size` bytes. An empty file needs one (empty) request. */
export function chunkCount(size, chunk = CHUNK_SIZE) {
  if (!(size > 0)) return 1;
  return Math.ceil(size / chunk);
}

/** Byte range for the chunk that starts at `offset`. `end` is exclusive. */
export function nextRange(offset, size, chunk = CHUNK_SIZE) {
  const start = Math.max(0, Math.min(offset, size));
  const end = Math.min(start + chunk, size);
  const header = size === 0 ? 'bytes */0' : `bytes ${start}-${end - 1}/${size}`;
  return { start, end, length: end - start, contentRange: header };
}

/** Header for "where are you?" queries on a resumable session. */
export function statusRangeHeader(size) { return `bytes */${size}`; }

/** Parse the Range response header of a 308 reply ("bytes=0-8388607") into the next offset. */
export function parseRangeHeader(value) {
  if (!value) return 0;
  const m = /bytes=\d+-(\d+)/i.exec(value);
  return m ? Number(m[1]) + 1 : 0;
}

/** Key under which a resumable session URI is remembered: destination + name + size + mtime. */
export function resumeKey(file, parentId, relPath) {
  return ['share-resume', parentId || '', relPath || file.name, file.size, file.lastModified || 0].join('|');
}

/** Exponential backoff with an upper bound (ms). `rand` in [0,1) adds jitter. */
export function backoffDelay(attempt, rand = 0.5, base = 800, cap = 30000) {
  const raw = Math.min(cap, base * Math.pow(2, Math.max(0, attempt - 1)));
  return Math.round(raw * (0.75 + rand * 0.5));
}

/** HTTP statuses worth retrying. */
export function isRetryable(status) {
  return status === 0 || status === 408 || status === 429 || (status >= 500 && status <= 599);
}

const EMAIL_RE = /^[^\s@<>()[\]\\,;:"]+@[^\s@<>()[\]\\,;:"]+\.[^\s@<>()[\]\\,;:"]{2,}$/;
export function isValidEmail(s) {
  return typeof s === 'string' && s.length <= 254 && EMAIL_RE.test(s.trim());
}

/** Split free text ("a@b.fr, c@d.com\n e@f.kr") into valid / invalid / duplicate lists. */
export function parseEmails(text) {
  const parts = String(text || '').split(/[\s,;]+/).map(s => s.trim()).filter(Boolean);
  const seen = new Set();
  const valid = [], invalid = [], duplicates = [];
  for (const p of parts) {
    const clean = p.replace(/^<|>$/g, '');
    if (!isValidEmail(clean)) { invalid.push(p); continue; }
    const k = clean.toLowerCase();
    if (seen.has(k)) { duplicates.push(clean); continue; }
    seen.add(k); valid.push(clean);
  }
  return { valid, invalid, duplicates };
}

/** Escape a value for a Drive `q` string literal. */
export function escapeQuery(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

/** Tidy a single file / folder / project name. Strips control characters and every character that
    OneDrive (and Windows) reserve: \ / : * ? " < > |  -> "-". Trailing dots/spaces and "."/".." are not valid either.
    The UI keeps showing the original name; only the stored name is cleaned. */
export function safeName(s) {
  let n = String(s || '').replace(/[\u0000-\u001f]/g, '').replace(/[\\/:*?"<>|]/g, '-').trim();
  n = n.slice(0, 200).replace(/[. ]+$/, '');
  if (n === '' && String(s || '').trim() !== '') n = '_';
  if (/^\.+$/.test(n)) n = '_';
  return n;
}
/** Clean every segment of a relative path ("a:b/c?.txt" -> "a-b/c-.txt"); empty segments are dropped. */
export function safePath(p) {
  return String(p || '').split('/').filter(x => x !== '').map(safeName).filter(Boolean).join('/');
}

/** ?mock=1 is honoured on a developer machine only. */
export function isLocalHost(hostname) {
  const h = String(hostname || '').toLowerCase();
  return h === 'localhost' || h === '127.0.0.1' || h === '[::1]' || h === '::1' || h.endsWith('.localhost');
}
/** href safety: only absolute https URLs from the storage provider are used as links. '' otherwise. */
export function safeHttpsUrl(u) {
  if (typeof u !== 'string') return '';
  try { const x = new URL(u.trim()); return x.protocol === 'https:' ? x.href : ''; } catch (e) { return ''; }
}
/** True only for the MSAL sign-in popup returning with its auth response (opener + hash + msal window name). */
export function isAuthPopup({ opener, hash, name } = {}) {
  return !!opener && /[#&](code|error)=/.test(String(hash || '')) && /^msal[._]/i.test(String(name || ''));
}

/**
 * entries: [{ path: 'a/b/c.txt', ...anything }]
 * Returns folder paths that must exist, parents first ("a", "a/b"), no duplicates.
 */
export function folderPaths(entries) {
  const set = new Set();
  for (const e of entries) {
    const parts = e.path.split('/').filter(Boolean);
    parts.pop();
    let acc = '';
    for (const p of parts) { acc = acc ? acc + '/' + p : p; set.add(acc); }
  }
  return [...set].sort((a, b) => a.split('/').length - b.split('/').length || a.localeCompare(b));
}

/** Split an entry path into { dir, name }. */
export function splitPath(path) {
  const i = path.lastIndexOf('/');
  return i < 0 ? { dir: '', name: path } : { dir: path.slice(0, i), name: path.slice(i + 1) };
}

/** Flat Drive listing → nested tree. items: [{id,name,mimeType,parents:[id],size}] */
export function buildFileTree(rootId, items) {
  const byParent = new Map();
  for (const it of items) {
    const p = (it.parents && it.parents[0]) || rootId;
    if (!byParent.has(p)) byParent.set(p, []);
    byParent.get(p).push(it);
  }
  const make = (id) => {
    const kids = (byParent.get(id) || []).map(it => {
      const isFolder = it.mimeType === FOLDER_MIME;
      return isFolder ? { ...it, isFolder, children: make(it.id) } : { ...it, isFolder };
    });
    kids.sort((a, b) => (b.isFolder - a.isFolder) || a.name.localeCompare(b.name, undefined, { numeric: true }));
    return kids;
  };
  return make(rootId);
}

/** Total bytes of every file in a tree. */
export function treeSize(nodes) {
  let n = 0;
  for (const x of nodes) n += x.isFolder ? treeSize(x.children || []) : Number(x.size || 0);
  return n;
}
export function treeCount(nodes) {
  let n = 0;
  for (const x of nodes) n += x.isFolder ? treeCount(x.children || []) : 1;
  return n;
}

export function formatBytes(n) {
  n = Number(n) || 0;
  if (n < 1024) return n + ' B';
  const u = ['KB', 'MB', 'GB', 'TB'];
  let i = -1;
  do { n /= 1024; i++; } while (n >= 1024 && i < u.length - 1);
  return (n >= 100 ? n.toFixed(0) : n >= 10 ? n.toFixed(1) : n.toFixed(2)) + ' ' + u[i];
}
export function formatSpeed(bps) { return bps > 0 ? formatBytes(bps) + '/s' : '–'; }
export function formatEta(sec) {
  if (!isFinite(sec) || sec < 0) return '–';
  sec = Math.round(sec);
  if (sec < 60) return sec + 's';
  const m = Math.floor(sec / 60), s = sec % 60;
  if (m < 60) return m + 'm ' + String(s).padStart(2, '0') + 's';
  const h = Math.floor(m / 60);
  return h + 'h ' + String(m % 60).padStart(2, '0') + 'm';
}

/** Rolling speed estimate. samples: [{t(ms), b(bytes uploaded)}] kept within `windowMs`. */
export function speedFromSamples(samples, windowMs = 6000) {
  if (samples.length < 2) return 0;
  const last = samples[samples.length - 1];
  let first = samples[0];
  for (const s of samples) if (last.t - s.t <= windowMs) { first = s; break; }
  const dt = (last.t - first.t) / 1000;
  return dt > 0 ? Math.max(0, (last.b - first.b) / dt) : 0;
}

/** Project URL in Drive. */
export function folderLink(id) { return 'https://drive.google.com/drive/folders/' + encodeURIComponent(id); }
export function fileLink(id) { return 'https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view'; }

/* ---------- OneDrive / Microsoft Graph ---------- */
export const GRAPH_ALIGN = 320 * 1024;            // upload fragments must be multiples of 320 KiB
export const ONEDRIVE_CHUNK = 32 * GRAPH_ALIGN;   // 10 MiB
export function isGraphAligned(n) { return n > 0 && n % GRAPH_ALIGN === 0; }
/** Largest 320 KiB multiple <= n (at least one unit). */
export function alignDown(n) { return Math.max(GRAPH_ALIGN, Math.floor(n / GRAPH_ALIGN) * GRAPH_ALIGN); }
/** nextExpectedRanges (["26-", "100-200"]) → first missing byte. */
export function parseNextExpected(ranges) {
  if (!Array.isArray(ranges) || !ranges.length) return 0;
  const n = Number(String(ranges[0]).split('-')[0]);
  return Number.isFinite(n) ? n : 0;
}
/** True for errors meaning "the organisation has not allowed this app". */
export function isAdminConsentError(text) {
  return /AADSTS(65001|65002|90094|900941|650056|700016)|admin_?consent|need admin approval|consent_required|interaction_required.*consent/i.test(String(text || ''));
}
/** Graph permission → common shape. */
export function mapGraphPermission(p) {
  const role = (p.roles && p.roles[0]) || 'read';
  const roleName = role === 'owner' ? 'owner' : role === 'write' ? 'writer' : 'reader';
  if (p.link) {
    const scope = p.link.scope;
    return { id: p.id, type: scope === 'anonymous' ? 'anyone' : scope === 'organization' ? 'organization' : 'link', role: roleName, url: p.link.webUrl };
  }
  const u = (p.grantedToV2 && (p.grantedToV2.user || p.grantedToV2.siteUser)) || (p.grantedTo && p.grantedTo.user) || {};
  const email = (p.invitation && p.invitation.email) || u.email || u.loginName || '';
  return { id: p.id, type: 'user', role: roleName, emailAddress: email, displayName: u.displayName || email };
}
