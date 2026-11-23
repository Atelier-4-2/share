/* Resumable upload task with pause / resume / retry. */
import { CHUNK_SIZE, MAX_RETRIES, nextRange, resumeKey, backoffDelay, isRetryable, speedFromSamples } from './lib.js';

const sleep = (ms, signal) => new Promise((res) => {
  const t = setTimeout(res, ms);
  if (signal) signal.addEventListener('abort', () => { clearTimeout(t); res(); }, { once: true });
});

/** Resume URIs live in localStorage; when it is blocked (Safari private mode, policies) an in-memory map keeps
    resume working for as long as this page lives. The memory copy is always written, so a failed or quota-full
    localStorage never loses the session URI. */
const mem = new Map();
export function store() {
  return {
    get(k) { let v = null; try { v = localStorage.getItem(k); } catch (e) { /* blocked */ } return v != null ? v : (mem.has(k) ? mem.get(k) : null); },
    set(k, v) { mem.set(k, v); try { localStorage.setItem(k, v); } catch (e) { /* blocked: memory copy only */ } },
    del(k) { mem.delete(k); try { localStorage.removeItem(k); } catch (e) { /* blocked */ } },
  };
}
const isAuthError = (e) => !!e && (e.code === 'auth-expired' || e.code === 'auth-required' || e.status === 401);

const splitName = (relPath, fallback) => { const i = String(relPath || '').lastIndexOf('/'); return (i < 0 ? relPath : relPath.slice(i + 1)) || fallback; };
let seq = 0;
export class UploadTask {
  /** opts: { drive, file, parentId, relPath (cleaned storage path), label (original name shown in the UI), overwrite, chunk, onUpdate, retryBase } */
  constructor(opts) {
    Object.assign(this, { retryBase: 800 }, opts);
    this.chunk = opts.chunk || (opts.drive && opts.drive.chunkSize) || CHUNK_SIZE;
    this.id = ++seq;
    this.size = this.file.size;
    this.state = 'queued'; // queued | uploading | retrying | paused | done | error | canceled
    this.sent = 0;
    this.attempt = 0;
    this.error = null;
    this.result = null;
    this.samples = [];
    this.key = resumeKey(this.file, this.parentId, this.relPath);
    this.kv = store();
    this.ctrl = null;
    this.running = false;
    this.resumedFromStore = false;
    this.needsAuth = false; // stopped because the sign-in expired: resumes after a re-login, no blind retries
  }
  get speed() { return speedFromSamples(this.samples); }
  get eta() { const s = this.speed; return s > 0 ? (this.size - this.sent) / s : Infinity; }
  _set(state) { this.state = state; this.onUpdate && this.onUpdate(this); }
  _progress(n) {
    this.sent = n; this.samples.push({ t: performance.now(), b: n });
    if (this.samples.length > 40) this.samples.shift();
    this.onUpdate && this.onUpdate(this);
  }

  start() {
    if (this.running || ['done', 'canceled'].includes(this.state)) return;
    this.running = true; this.ctrl = new AbortController();
    this.needsAuth = false;
    this._run().catch((e) => { this.error = e; this.needsAuth = isAuthError(e); this._set('error'); }).finally(() => { this.running = false; });
  }
  pause() { if (this.state === 'uploading' || this.state === 'retrying' || this.state === 'queued') { this.ctrl && this.ctrl.abort(); this._set('paused'); } }
  resume() { if (this.state === 'paused' || this.state === 'error') { this.error = null; this.needsAuth = false; this.attempt = 0; this.start(); } }
  cancel() { this.ctrl && this.ctrl.abort(); this.kv.del(this.key); this._set('canceled'); }

  async _openSession(signal) {
    let uri = this.kv.get(this.key);
    // a stored URI is only sent a token if the provider would have handed out exactly such an address
    if (uri && this.drive.isValidUploadUri && !this.drive.isValidUploadUri(uri)) { this.kv.del(this.key); uri = null; }
    if (uri) {
      try {
        const q = await this.drive.queryUpload(uri, this.size, signal);
        if (q.done) { this.kv.del(this.key); return { uri, done: q.file }; }
        this.resumedFromStore = true;
        return { uri, offset: q.next };
      } catch (e) {
        if (e && e.name === 'AbortError') throw e;
        // sign-in expired or the network blinked: keep the session URI, the caller waits / retries and resumes it
        if (isAuthError(e) || isRetryable(e && e.status)) throw e;
        // 404/410/400 = session gone: fall back to a fresh session
        this.kv.del(this.key);
      }
    }
    uri = await this.drive.startUpload({ name: this.storeName || splitName(this.relPath, this.file.name), parentId: this.parentId, size: this.size, mimeType: this.file.type, overwrite: !!this.overwrite });
    this.kv.set(this.key, uri);
    return { uri, offset: 0 };
  }

  async _run() {
    const signal = this.ctrl.signal;
    this._set('uploading');
    let session;
    for (let a = 1; ; a++) {
      try { session = await this._openSession(signal); break; }
      catch (e) {
        if (signal.aborted) return;
        if (a >= MAX_RETRIES || !(isRetryable(e.status))) throw e;
        this._set('retrying'); await sleep(backoffDelay(a, Math.random(), this.retryBase), signal);
        if (signal.aborted) return;
      }
    }
    if (session.done) { this.result = session.done; this._progress(this.size); return this._set('done'); }
    let { uri, offset } = session;
    this._progress(offset);
    this._set('uploading');

    outer: while (true) {
      if (signal.aborted) return;
      const r = nextRange(offset, this.size, this.chunk);
      let res;
      for (this.attempt = 1; ; this.attempt++) {
        try {
          res = await this.drive.uploadChunk(uri, this.file.slice(r.start, r.end), r.start, this.size, signal);
          if (this.state === 'retrying') this._set('uploading');
          break;
        } catch (e) {
          if (signal.aborted) return;
          const expired = e && (e.status === 404 || e.status === 410);
          if (expired) { this.kv.del(this.key); throw e; }
          if (!isRetryable(e && e.status) || this.attempt >= MAX_RETRIES) throw e;
          this._set('retrying');
          await sleep(backoffDelay(this.attempt, Math.random(), this.retryBase), signal);
          if (signal.aborted) return;
          // find out what Google actually received before sending again
          try {
            const q = await this.drive.queryUpload(uri, this.size, signal);
            if (q.done) { res = q; break; }
            if (q.next !== r.start) { offset = q.next; this._progress(offset); }
          } catch (e2) { if (signal.aborted) return; }
          if (res) break;
          if (offset !== r.start) continue outer; // server has a different offset: re-slice from there, same session
        }
      }
      if (res.done) { this.kv.del(this.key); this.result = res.file; this._progress(this.size); return this._set('done'); }
      offset = res.next; this._progress(offset);
    }
  }
}

/** Runs up to `limit` tasks at once. */
export class UploadQueue {
  constructor(limit = 2) { this.limit = limit; this.tasks = []; }
  add(task) { this.tasks.push(task); this.pump(); return task; }
  pump() {
    const active = this.tasks.filter(t => t.running).length;
    let free = this.limit - active;
    for (const t of this.tasks) {
      if (free <= 0) break;
      if (t.state === 'queued' && !t.running) { t.start(); free--; }
    }
  }
}
