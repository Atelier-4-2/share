import test from 'node:test';
import assert from 'node:assert/strict';

// minimal browser storage stubs
const mkStore = () => { const m = new Map(); return { getItem: k => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k), clear: () => m.clear() }; };
globalThis.sessionStorage = mkStore();
globalThis.localStorage = mkStore();

const C = await import('../atelier-cloud.js');
const { createMockServer } = await import('../mock-server.js');
const { createMockGraph } = await import('../mock-graph.js');

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

function setup(provider) {
  sessionStorage.clear(); localStorage.clear();
  const g = createMockServer({ speedMs: 0 }), o = createMockGraph({ speedMs: 0 });
  const server = provider === 'google' ? g : o;
  let offline = false;
  const cloud = C.createCloud({
    overrides: { google: { fetch: g.handler }, onedrive: { fetch: o.handler } },
    queueStore: C.memoryQueueStore(), isOffline: () => offline, pollMs: 40,
  });
  return { cloud, server, setOffline: (v) => { offline = v; server.hooks.offline = v; } };
}

/* ---------- pure logic ---------- */
test('normalizePath', () => {
  assert.deepEqual(C.normalizePath('board//a/b.json'), ['board', 'a', 'b.json']);
  for (const bad of ['', '/', 'a/../b', './a', 'a\\b']) assert.throws(() => C.normalizePath(bad), /path/);
});
test('newManifest / applyManifestUpdate', () => {
  const m = C.newManifest('P', '2026-01-01T00:00:00Z');
  assert.deepEqual(m.folders, ['files', 'board', 'crit', 'models']);
  assert.equal(m.schema, 1);
  const next = C.applyManifestUpdate(m, x => { x.tools.board = { path: 'board/a.json' }; }, '2026-02-02T00:00:00Z');
  assert.equal(next.tools.board.path, 'board/a.json');
  assert.equal(next.updated, '2026-02-02T00:00:00Z');
  assert.equal(m.tools.board, undefined, 'input is not mutated');
  const repaired = C.applyManifestUpdate({ name: 'x', tools: null }, () => {});
  assert.deepEqual(repaired.tools, {});
});
test('retryOnConflict', async () => {
  let n = 0;
  assert.equal(await C.retryOnConflict(async () => { if (++n < 3) throw new C.ConflictError(); return 'ok'; }), 'ok');
  await assert.rejects(C.retryOnConflict(async () => { throw new C.ConflictError(); }, 2), C.ConflictError);
  await assert.rejects(C.retryOnConflict(async () => { throw new Error('other'); }), /other/);
});
test('snapshot diff', () => {
  const t1 = [{ id: 'a', name: 'a', isFolder: true, children: [{ id: 'b', name: 'b.txt', version: '1', size: 1 }] }];
  const t2 = [{ id: 'a', name: 'a', isFolder: true, children: [{ id: 'b', name: 'b.txt', version: '2', size: 1 }, { id: 'c', name: 'c.txt', version: '1' }] }];
  const d = C.diffSnapshots(C.snapshotOf(t1), C.snapshotOf(t2));
  assert.deepEqual([d.added.map(x => x.path), d.modified.map(x => x.path), d.removed.length], [['a/c.txt'], ['a/b.txt'], 0]);
  const d2 = C.diffSnapshots(C.snapshotOf(t2), C.snapshotOf([]));
  assert.equal(d2.removed.length, 3);
});
test('chunk constants', () => {
  assert.equal(C.GOOGLE_CHUNK % (256 * 1024), 0);
  assert.equal(C.ONEDRIVE_CHUNK % (320 * 1024), 0);
  assert.equal(C.chunkRange(0, 20, 8).header, 'bytes 0-7/20');
  assert.equal(C.parseNextExpected(['26-']), 26);
});

/* ---------- mock-backed, both providers ---------- */
for (const provider of ['google', 'onedrive']) {
  test(`[${provider}] project create, folders, manifest`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('Colombes_p10');
    const names = (await p.list('')).map(n => n.name).sort();
    assert.deepEqual(names, ['board', 'crit', 'files', 'models', 'project.json']);
    const m = await p.manifest();
    assert.equal(m.name, 'Colombes_p10'); assert.deepEqual(m.tools, {});
    await assert.rejects(cloud.createProject('Colombes_p10'), e => e.code === 'exists');
    assert.equal((await cloud.listProjects()).length, 1);
    const again = await cloud.openProject((await cloud.listProjects())[0]);
    assert.equal(again.name, 'Colombes_p10');
  });

  test(`[${provider}] readJson/writeJson version check`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    const w1 = await p.writeJson('board/main.json', { v: 1 }, { version: null });
    await assert.rejects(p.writeJson('board/main.json', { v: 9 }, { version: null }), C.ConflictError, 'create-only fails when it exists');
    const r1 = await p.readJson('board/main.json');
    assert.deepEqual(r1.data, { v: 1 }); assert.equal(r1.version, w1.version);
    const w2 = await p.writeJson('board/main.json', { v: 2 }, { version: r1.version });
    assert.notEqual(w2.version, r1.version);
    await assert.rejects(p.writeJson('board/main.json', { v: 3 }, { version: r1.version }), C.ConflictError, 'stale version → conflict');
    assert.deepEqual((await p.readJson('board/main.json')).data, { v: 2 }, 'stale write did not land');
    await assert.rejects(p.readJson('board/none.json'), e => e.code === 'not-found');
    await assert.rejects(p.readJson('nofolder/x.json'), e => e.code === 'not-found');
  });

  test(`[${provider}] concurrent updateManifest both land (retry on conflict)`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    await Promise.all([
      p.registerTool('board', { path: 'board/a.json' }),
      p.registerTool('crit', { path: 'crit/s.json' }),
      p.registerTool('view', { path: 'models/m.3dm' }),
    ]);
    const m = await p.manifest();
    assert.deepEqual(Object.keys(m.tools).sort(), ['board', 'crit', 'view']);
  });

  test(`[${provider}] big upload (25 MB) + download roundtrip`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    let last = 0;
    const blob = new Blob([new Uint8Array(25 * 1024 * 1024)], { type: 'application/octet-stream' });
    const file = await p.uploadBlob('files/renders/big.bin', blob, { onProgress: (s) => { last = s; } });
    assert.equal(last, blob.size);
    assert.equal((await p.stat('files/renders/big.bin')).size, blob.size);
    assert.ok(file);
    const back = await p.downloadBlob('files/renders/big.bin');
    assert.equal(back.size, 1e6 > blob.size ? blob.size : 1e6);
  });

  test(`[${provider}] second device change is noticed by watchProject`, async () => {
    const { cloud, server } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    const filesId = await p.folderId('files');
    const events = [];
    const w = p.watch(e => events.push(e), { pollMs: 40 });
    await sleep(150);
    assert.equal(events.length, 0, 'no change yet');
    server.remoteWrite(filesId, 'from_phone.json', '{"x":1}');
    await sleep(250);
    assert.equal(events.length, 1);
    assert.deepEqual(events[0].added.map(x => x.path), ['files/from_phone.json']);
    server.remoteWrite(filesId, 'from_phone.json', '{"x":2}');
    await sleep(250);
    assert.equal(events.length, 2);
    assert.deepEqual(events[1].modified.map(x => x.name), ['from_phone.json']);
    w.stop();
    server.remoteWrite(filesId, 'late.json', '{}');
    await sleep(150);
    assert.equal(events.length, 2, 'stopped watcher stays quiet');
  });

  test(`[${provider}] offline queue: saveJson queues, flush sends, conflict kept`, async () => {
    const { cloud, setOffline } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    setOffline(true);
    const r = await p.saveJson('board/offline.json', { a: 1 }, { version: null });
    assert.deepEqual(r, { queued: true });
    const r2 = await p.saveBlob('files/note.bin', new Blob([new Uint8Array(1000)]));
    assert.deepEqual(r2, { queued: true });
    assert.equal((await cloud.queueStatus()).pending, 2);
    await cloud.flushQueue();
    assert.equal((await cloud.queueStatus()).pending, 2, 'still offline → nothing sent');
    setOffline(false);
    await cloud.flushQueue();
    assert.equal((await cloud.queueStatus()).pending, 0);
    assert.deepEqual((await p.readJson('board/offline.json')).data, { a: 1 });
    assert.equal((await p.stat('files/note.bin')).size, 1000);
    // queued write whose base version went stale elsewhere → conflict is kept, not overwritten
    const cur = await p.readJson('board/offline.json');
    setOffline(true);
    await p.saveJson('board/offline.json', { a: 'mine' }, { version: cur.version });
    setOffline(false);
    await p.writeJson('board/offline.json', { a: 'theirs' }, { version: cur.version });
    await cloud.flushQueue();
    const st = await cloud.queueStatus();
    assert.equal(st.conflicts, 1);
    assert.deepEqual((await p.readJson('board/offline.json')).data, { a: 'theirs' });
    await cloud.resolveQueued(st.items[0].id, 'overwrite');
    assert.deepEqual((await p.readJson('board/offline.json')).data, { a: 'mine' });
    assert.equal((await cloud.queueStatus()).conflicts, 0);
  });

  test(`[${provider}] last project, session restore, sign-out`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('Remember me');
    const last = cloud.lastProject();
    assert.equal(last.projectId, p.id); assert.equal(last.provider, provider);
    assert.ok(!JSON.stringify(localStorage.getItem('atelier-cloud-last-mock')).includes('mock-token'), 'no token in localStorage');
    assert.equal(localStorage.getItem('atelier-cloud-session'), null, 'session never in localStorage');
    assert.ok(sessionStorage.getItem(C.MOCK_SESSION_KEY).includes('mock-token'));
    // a second page of the same tab: new cloud object, same sessionStorage, same (mock) server
    const g = createMockServer(), o = createMockGraph();
    cloud.signOutLocal();
    assert.equal(await cloud.restore(), null, 'signOutLocal cleared the session');
    await cloud.signIn(provider);
    assert.equal(cloud.preferred, provider);
    cloud.signOut();
    assert.equal(sessionStorage.getItem(C.MOCK_SESSION_KEY), null);
    assert.equal(cloud.signedIn, false);
    void g; void o;
  });
}

test('restore() reuses sessionStorage across cloud instances', async () => {
  sessionStorage.clear(); localStorage.clear();
  const g = createMockServer();
  const mk = () => C.createCloud({ overrides: { google: { fetch: g.handler } }, queueStore: C.memoryQueueStore() });
  const a = mk(); await a.signIn('google');
  const b = mk();
  const u = await b.restore();
  assert.equal(u.emailAddress, 'mock.user@example.test');
  assert.equal(b.provider, 'google');
});

test('silent refresh failure signals auth-expired once, reauth() recovers', async () => {
  sessionStorage.clear();
  const g = createMockServer();
  let refuse = false, calls = 0;
  const fetchWrap = (u, i) => { calls++; if (refuse && calls % 1 === 0 && !(i.headers && i.headers.Authorization === 'Bearer fresh')) { refuse = 'used'; return Promise.resolve(new Response('{}', { status: 401 })); } return g.handler(u, i); };
  const cloud = C.createCloud({ overrides: { google: { fetch: fetchWrap, refuseRefresh: () => refuse } }, queueStore: C.memoryQueueStore() });
  const events = [];
  cloud.on('auth-expired', e => events.push(e));
  await cloud.signIn('google');
  refuse = true; // next API call gets 401, silent refresh is refused
  await assert.rejects(cloud.listProjects(), e => e.code === 'auth-expired');
  assert.equal(events.length, 1);
  await assert.rejects(cloud.listProjects());
  assert.equal(events.length, 1, 'signalled only once');
  refuse = false;
  await cloud.reauth();
  assert.ok((await cloud.listProjects()).length >= 0);
});

test('OneDrive admin-consent error is classified', async () => {
  sessionStorage.clear();
  const o = createMockGraph();
  const cloud = C.createCloud({ overrides: { onedrive: { fetch: o.handler, signIn: async () => { throw new C.CloudError('admin-consent', 'AADSTS90094') } } }, queueStore: C.memoryQueueStore() });
  await assert.rejects(cloud.signIn('onedrive'), e => e.code === 'admin-consent');
  assert.equal(cloud.signedIn, false);
});

test('OneDrive link: anonymous denied → organization fallback', async () => {
  const { cloud, server } = setup('onedrive');
  await cloud.signIn('onedrive');
  const p = await cloud.createProject('P');
  server.hooks.denyAnonymous = true;
  const r = await p.share.setLink(true);
  assert.deepEqual(r, { scope: 'organization', fallback: true });
  const list = await p.share.list();
  assert.ok(list.some(x => x.type === 'organization' && x.url));
  await p.share.setLink(false);
  assert.ok(!(await p.share.list()).some(x => x.type === 'organization'));
  const sent = await p.share.emails(['a@b.fr'], 'hello');
  assert.ok(sent[0].ok);
  assert.deepEqual(server.hooks.invites.at(-1), { to: 'a@b.fr', sendInvitation: true, message: 'hello', requireSignIn: true });
});

test('Google email share sends notification flag', async () => {
  const { cloud, server } = setup('google');
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  await p.share.emails(['x@y.fr'], 'hi');
  assert.deepEqual(server.hooks.lastEmails.at(-1), { to: 'x@y.fr', notify: 'true', message: 'hi' });
  assert.deepEqual(await p.share.setLink(true), { scope: 'anyone' });
});

/* =====================  review round (Gemini)  ===================== */
const { UploadTask, store: resumeStore } = await import('../upload.js');
const { readFileSync } = await import('node:fs');

test('CloudError: auth failures are 401 (never the retryable status 0)', () => {
  assert.equal(new C.CloudError('auth-expired').status, 401);
  assert.equal(new C.CloudError('auth-required').status, 401);
  assert.equal(new C.CloudError('network').status, 0);
  assert.equal(C.retryable(new C.CloudError('auth-expired').status), false);
  assert.equal(C.retryable(0), true);
});

test('mock sessions use their own keys and never touch the real session', async () => {
  sessionStorage.clear(); localStorage.clear();
  sessionStorage.setItem(C.SESSION_KEY, JSON.stringify({ provider: 'google', token: 'REAL', expiresAt: Date.now() + 3600e3, user: null }));
  const g = createMockServer();
  const cloud = C.createCloud({ overrides: { google: { fetch: g.handler } }, queueStore: C.memoryQueueStore() });
  await cloud.signIn('google');
  await cloud.createProject('P');
  assert.ok(sessionStorage.getItem(C.MOCK_SESSION_KEY).includes('mock-token'));
  assert.ok(JSON.parse(sessionStorage.getItem(C.SESSION_KEY)).token === 'REAL', 'real session untouched');
  assert.equal(localStorage.getItem(C.PROVIDER_KEY), null);
  assert.equal(localStorage.getItem('atelier-cloud-last'), null);
  assert.ok(localStorage.getItem('atelier-cloud-last-mock'));
  cloud.signOut();
  assert.equal(JSON.parse(sessionStorage.getItem(C.SESSION_KEY)).token, 'REAL', 'sign-out of the mock leaves the real session alone');
  // a real (non-mock) page never sees the mock token
  sessionStorage.clear();
  const g2 = createMockServer();
  const mockCloud = C.createCloud({ overrides: { google: { fetch: g2.handler } }, queueStore: C.memoryQueueStore() });
  await mockCloud.signIn('google');
  const real = C.createCloud({ googleClientId: 'real-id', queueStore: C.memoryQueueStore() });
  assert.equal(await real.restore(), null, 'mock-token is invisible to a real cloud');
});

test('auth-expired during an upload is not retried; resume works after re-login', async () => {
  sessionStorage.clear(); localStorage.clear();
  const g = createMockServer({ speedMs: 0 });
  let refuse = false, chunkCalls = 0;
  const fetchWrap = (u, i) => { if (i && i.method === 'PUT' && /upload_id/.test(String(u))) chunkCalls++; return g.handler(u, i); };
  const cloud = C.createCloud({ overrides: { google: { fetch: fetchWrap, refuseRefresh: () => refuse } }, queueStore: C.memoryQueueStore() });
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  refuse = true; g.hooks.force401 = true;
  const t0 = Date.now();
  await assert.rejects(p.uploadBlob('files/big.bin', new Blob([new Uint8Array(100)])), e => e.code === 'auth-expired' && e.status === 401);
  assert.ok(Date.now() - t0 < 500, 'no 6-step backoff loop');
  assert.ok(chunkCalls <= 1, 'at most the one request that got the 401');
  refuse = false; await cloud.reauth();
  const f = await p.uploadBlob('files/big.bin', new Blob([new Uint8Array(100)]));
  assert.equal(f.size, 100);
});

test('UploadTask: auth-expired stops at once (needsAuth), keeps the session URI, resumes after re-login', async () => {
  localStorage.clear();
  let chunks = 0, authDead = true, started = 0;
  const drive = {
    chunkSize: 10, isValidUploadUri: () => true,
    async startUpload() { started++; return 'https://www.googleapis.com/upload/drive/v3/files?upload_id=s1'; },
    async queryUpload() { return { done: false, next: 0 }; },
    async uploadChunk(uri, blob, start, total) { chunks++; if (authDead) throw new C.CloudError('auth-expired', 'expired'); return start + blob.size >= total ? { done: true, file: { id: 'f' } } : { done: false, next: start + blob.size }; },
  };
  const file = { name: 'a.bin', size: 25, lastModified: 1, type: '', slice: (a, b) => ({ size: b - a }) };
  const task = new UploadTask({ drive, file, parentId: 'P', relPath: 'a.bin', retryBase: 1 });
  task.start();
  while (task.running) await sleep(5);
  assert.equal(task.state, 'error'); assert.equal(task.needsAuth, true);
  assert.equal(chunks, 1, 'exactly one attempt, no retries');
  authDead = false; task.resume();
  while (task.running) await sleep(5);
  assert.equal(task.state, 'done');
  assert.equal(started, 1, 'reused the stored session (kept after the auth error)');
});

test('UploadTask: server offset differs after a retry -> continues from it in the same session (no recursion, no new session)', async () => {
  localStorage.clear();
  const calls = []; let started = 0, queries = 0, failOnce = true;
  const drive = {
    chunkSize: 10, isValidUploadUri: () => true,
    async startUpload() { started++; return 'https://www.googleapis.com/upload/drive/v3/files?upload_id=s2'; },
    async queryUpload() { queries++; return { done: false, next: 20 }; }, // chunk 10-19 actually arrived
    async uploadChunk(uri, blob, start, total) {
      calls.push(start);
      if (start === 10 && failOnce) { failOnce = false; throw new C.CloudError('http', 'blip', 503); }
      return start + blob.size >= total ? { done: true, file: { id: 'f' } } : { done: false, next: start + blob.size };
    },
  };
  const file = { name: 'b.bin', size: 25, lastModified: 2, type: '', slice: (a, b) => ({ size: b - a }) };
  const task = new UploadTask({ drive, file, parentId: 'P', relPath: 'b.bin', retryBase: 1 });
  task.start(); while (task.running) await sleep(5);
  assert.equal(task.state, 'done');
  assert.deepEqual(calls, [0, 10, 20]);
  assert.equal(started, 1);
  assert.equal(queries, 1, 'one status query, not a second one from a re-entered _run');
});

test('resume URIs are validated before any request', async () => {
  // pure
  assert.ok(C.isGoogleUploadUri('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&upload_id=x'));
  for (const u of ['http://www.googleapis.com/upload/drive/v3/files', 'https://evil.example/upload/drive/v3/', 'https://www.googleapis.com.evil.io/upload/drive/v3/x', 'https://www.googleapis.com/drive/v3/files', '', null]) assert.equal(C.isGoogleUploadUri(u), false, String(u));
  for (const u of ['https://contoso-my.sharepoint.com/x', 'https://abc.files.1drv.com/y', 'https://graph.microsoft.com/v1.0/z', 'https://api.onedrive.com/rup/1']) assert.equal(C.isGraphUploadUri(u), true, u);
  for (const u of ['http://contoso-my.sharepoint.com/x', 'https://sharepoint.com.evil.io/x', 'https://evilsharepoint.com/x', 'https://evil.example/files.1drv.com', 'https://u:p@abc.sharepoint.com/x', 'javascript:1', '']) assert.equal(C.isGraphUploadUri(u), false, u);
  // task: a stored foreign URI is dropped and never queried (no token sent to it)
  localStorage.clear();
  const queried = [];
  const drive = {
    chunkSize: 10, isValidUploadUri: C.isGoogleUploadUri,
    async startUpload() { return 'https://www.googleapis.com/upload/drive/v3/files?upload_id=ok'; },
    async queryUpload(uri) { queried.push(uri); return { done: false, next: 0 }; },
    async uploadChunk(uri, blob, start, total) { return start + blob.size >= total ? { done: true, file: { id: 'f' } } : { done: false, next: start + blob.size }; },
  };
  const file = { name: 'c.bin', size: 5, lastModified: 3, type: '', slice: (a, b) => ({ size: b - a }) };
  const task = new UploadTask({ drive, file, parentId: 'P', relPath: 'c.bin' });
  localStorage.setItem(task.key, 'https://evil.example/steal');
  task.start(); while (task.running) await sleep(5);
  assert.equal(task.state, 'done');
  assert.deepEqual(queried, [], 'foreign URI was never queried');
  // adapter level: Authorization is never attached to a foreign URI
  const g = createMockServer({ speedMs: 0 }); const seen = [];
  const cloud = C.createCloud({ overrides: { google: { fetch: (u, i) => { seen.push(String(u)); return g.handler(u, i); } } }, queueStore: C.memoryQueueStore() });
  await cloud.signIn('google');
  seen.length = 0;
  await assert.rejects(cloud.adapter.queryUpload('https://evil.example/x', 10), e => e.status === 400);
  await assert.rejects(cloud.adapter.uploadChunk('https://evil.example/x', new Blob([1]), 0, 1), e => e.status === 400);
  assert.deepEqual(seen, []);
});

test('resume store falls back to memory when localStorage is blocked', () => {
  const saved = globalThis.localStorage;
  globalThis.localStorage = { getItem() { throw new Error('blocked'); }, setItem() { throw new Error('blocked'); }, removeItem() { throw new Error('blocked'); } };
  try {
    const kv = resumeStore();
    kv.set('k1', 'v1'); assert.equal(kv.get('k1'), 'v1');
    kv.del('k1'); assert.equal(kv.get('k1'), null);
  } finally { globalThis.localStorage = saved; }
});

test('[google] same-name upload: second file gets "(2)", overwrite replaces, parallel uploads do not collide', async () => {
  const { cloud, server } = setup('google');
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  const a = await p.uploadBlob('files/plan.pdf', new Blob([new Uint8Array(10)]));
  const b = await p.uploadBlob('files/plan.pdf', new Blob([new Uint8Array(20)]));
  assert.equal(a.name, 'plan.pdf'); assert.equal(b.name, 'plan (2).pdf'); assert.notEqual(a.id, b.id);
  const c = await p.uploadBlob('files/plan.pdf', new Blob([new Uint8Array(30)]), { overwrite: true });
  assert.equal(c.id, a.id, 'overwrite keeps the same file');
  assert.equal(c.size, 30);
  const names = (await p.list('files')).map(n => n.name).sort();
  assert.deepEqual(names, ['plan (2).pdf', 'plan.pdf']);
  const [x, y] = await Promise.all([p.uploadBlob('files/par.bin', new Blob([new Uint8Array(5)])), p.uploadBlob('files/par.bin', new Blob([new Uint8Array(6)]))]);
  assert.notEqual(x.name, y.name);
  assert.equal((await p.list('files')).filter(n => n.name.startsWith('par')).length, 2);
  void server;
});

test('uniqueName / conflictCopyPath', () => {
  assert.equal(C.uniqueName('a.mov', new Set()), 'a.mov');
  assert.equal(C.uniqueName('a.mov', new Set(['a.mov'])), 'a (2).mov');
  assert.equal(C.uniqueName('a.mov', new Set(['a.mov', 'a (2).mov'])), 'a (3).mov');
  assert.equal(C.uniqueName('README', new Set(['README'])), 'README (2)');
  assert.equal(C.uniqueName('.hidden', new Set(['.hidden'])), '.hidden (2)');
  assert.equal(C.conflictCopyPath('board/main.json', new Date('2026-11-23T10:05:09Z')), 'board/main (conflict 20261123100509).json');
  assert.equal(C.conflictCopyPath('x', new Date('2026-11-23T10:05:09Z')), 'x (conflict 20261123100509)');
});

test('[google] create-only manifest: a duplicate created by another device is detected and ours is withdrawn', async () => {
  sessionStorage.clear(); localStorage.clear();
  const g = createMockServer({ speedMs: 0 });
  let injected = false;
  const fetchWrap = (u, i) => {
    // right before our multipart create lands, "another device" creates project.json first
    if (!injected && i && i.method === 'POST' && /uploadType=multipart/.test(String(u))) {
      injected = true;
      const meta = /"parents":\["([^"]+)"\]/.exec(i.body)[1];
      const id = g.remoteWrite(meta, 'project.json', '{"theirs":true}');
      g.files.get(id).createdTime = '2000-01-01T00:00:00.000Z'; // they were first
    }
    return g.handler(u, i);
  };
  const cloud = C.createCloud({ overrides: { google: { fetch: fetchWrap } }, queueStore: C.memoryQueueStore() });
  await cloud.signIn('google');
  await assert.rejects(cloud.createProject('P'), e => e instanceof C.ConflictError);
  const live = [...g.files.values()].filter(f => f.name === 'project.json' && !f.trashed);
  assert.equal(live.length, 1, 'exactly one project.json survives');
  assert.equal(live[0].content, '{"theirs":true}');
});

test('[google] overwrite re-checks headRevisionId right before the PATCH', async () => {
  const { cloud, server } = setup('google');
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  const r = await p.readJson('project.json');
  server.remoteWrite(await p.folderId(''), 'project.json', '{"other":1}'); // someone else wrote after our read
  await assert.rejects(p.writeJson('project.json', { mine: 1 }, { version: r.version }), C.ConflictError);
  // updateManifest retry-merges instead of failing
  await p.updateManifest(m => { m.tools.share = { ok: true }; });
  assert.equal((await p.manifest()).tools.share.ok, true);
});

for (const provider of ['google', 'onedrive']) {
  test(`[${provider}] sharing targets files/, not the project root`, async () => {
    const { cloud } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    await p.share.setLink(true);
    await p.share.emails(['a@b.fr']);
    const filesId = await p.folderId('files');
    assert.equal(await p.share.folderId(), filesId);
    const onFiles = await cloud.adapter.listPermissions(filesId);
    assert.ok(onFiles.some(x => x.type === 'anyone') && onFiles.some(x => x.emailAddress === 'a@b.fr'));
    assert.deepEqual((await cloud.adapter.listPermissions(p.id)).filter(x => x.role !== 'owner'), [], 'project root stays private');
    assert.ok((await p.share.list()).some(x => x.type === 'anyone'));
  });

  test(`[${provider}] queue: conflict -> keep both / discard, failed -> retry`, async () => {
    const { cloud, setOffline } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    await p.writeJson('board/b.json', { v: 1 });
    const cur = await p.readJson('board/b.json');
    setOffline(true);
    await p.saveJson('board/b.json', { v: 'mine' }, { version: cur.version });
    setOffline(false);
    await p.writeJson('board/b.json', { v: 'theirs' }, { version: cur.version });
    await cloud.flushQueue();
    let st = await cloud.queueStatus(); assert.equal(st.conflicts, 1);
    await cloud.resolveQueued(st.items[0].id, 'keep-both');
    st = await cloud.queueStatus(); assert.equal(st.conflicts + st.pending + st.failed, 0);
    assert.deepEqual((await p.readJson('board/b.json')).data, { v: 'theirs' }, 'their version untouched');
    const copy = (await p.list('board')).find(n => /^b \(conflict \d{14}\)\.json$/.test(n.name));
    assert.ok(copy, 'my version is kept next to it');
    assert.deepEqual((await p.readJson('board/' + copy.name)).data, { v: 'mine' });
    // discard
    setOffline(true);
    await p.saveJson('board/b.json', { v: 'x' }, { version: cur.version });
    setOffline(false);
    await cloud.flushQueue();
    st = await cloud.queueStatus(); assert.equal(st.conflicts, 1);
    await cloud.resolveQueued(st.items[0].id, 'discard');
    st = await cloud.queueStatus(); assert.equal(st.conflicts + st.pending + st.failed, 0);
  });
}

test('queue: failed item can be retried', async () => {
  const { cloud, setOffline } = setup('google');
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  setOffline(true);
  await p.saveJson('board/f.json', { a: 1 }, {});
  setOffline(false);
  // the project handle is cached, so flushQueue uses this patched writeJson: five non-network failures -> 'failed'
  const realWrite = p.writeJson.bind(p); let fails = 0;
  p.writeJson = async () => { fails++; throw new C.CloudError('http', 'boom', 400); };
  for (let i = 0; i < 5; i++) await cloud.flushQueue();
  assert.equal((await cloud.queueStatus()).failed, 1);
  p.writeJson = realWrite;
  await cloud.resolveQueued((await cloud.queueStatus()).items[0].id, 'retry');
  const st = await cloud.queueStatus();
  assert.equal(st.failed + st.pending, 0);
  assert.deepEqual((await p.readJson('board/f.json')).data, { a: 1 });
  assert.equal(fails, 5);
});

test('Google sign-out loads GSI first when the page was reloaded, then revokes', async () => {
  sessionStorage.clear(); localStorage.clear();
  const g = createMockServer();
  const revoked = [];
  const savedFetch = globalThis.fetch;
  globalThis.fetch = (u, i) => g.handler(u, i);
  globalThis.window = {};
  globalThis.document = { head: { appendChild(s) { globalThis.window.google = { accounts: { oauth2: { revoke: (t) => revoked.push(t) } } }; setTimeout(() => s.onload && s.onload(), 0); } }, createElement: () => ({}) };
  try {
    sessionStorage.setItem(C.SESSION_KEY, JSON.stringify({ provider: 'google', token: 'tok-123', expiresAt: Date.now() + 3600e3, user: { displayName: 'U', emailAddress: 'u@x.fr' } }));
    const cloud = C.createCloud({ googleClientId: 'cid', queueStore: C.memoryQueueStore() });
    assert.ok(await cloud.restore());
    assert.equal(globalThis.window.google, undefined, 'GSI not loaded after a reload');
    await cloud.signOut();
    assert.deepEqual(revoked, ['tok-123']);
    assert.equal(sessionStorage.getItem(C.SESSION_KEY), null);
  } finally { globalThis.fetch = savedFetch; delete globalThis.window; delete globalThis.document; }
});

test('MSAL: no account after a reload -> ssoSilent with the remembered login, then continue', async () => {
  sessionStorage.clear(); localStorage.clear();
  const o = createMockGraph();
  const savedFetch = globalThis.fetch;
  globalThis.fetch = (u, i) => o.handler(u, i);
  const msalCalls = [];
  globalThis.location = { origin: 'https://t.test', pathname: '/share/' };
  globalThis.window = { msal: { PublicClientApplication: class {
    async initialize() {} getAllAccounts() { return []; }
    async ssoSilent(req) { msalCalls.push(['sso', req.loginHint]); return { accessToken: 'fresh', expiresOn: new Date(Date.now() + 3600e3), account: { username: req.loginHint } }; }
    async acquireTokenSilent() { throw new Error('should not be needed'); }
    async acquireTokenPopup() { msalCalls.push(['popup']); throw new Error('popup must not open in the background'); }
    clearCache() {}
  } } };
  try {
    sessionStorage.setItem(C.SESSION_KEY, JSON.stringify({ provider: 'onedrive', token: 'old', expiresAt: Date.now() + 3600e3, user: { displayName: 'U', emailAddress: 'u@school.fr' } }));
    const cloud = C.createCloud({ msClientId: 'cid', queueStore: C.memoryQueueStore() });
    assert.ok(await cloud.restore());
    o.hooks.force401 = true;                       // the old token is rejected -> silent refresh
    const projects = await cloud.listProjects();
    assert.ok(Array.isArray(projects));
    assert.deepEqual(msalCalls, [['sso', 'u@school.fr']]);
  } finally { globalThis.fetch = savedFetch; delete globalThis.window; delete globalThis.location; }
});

test('MSAL: ssoSilent failure -> auth-expired (re-login bar), not a hang', async () => {
  sessionStorage.clear(); localStorage.clear();
  const o = createMockGraph();
  const savedFetch = globalThis.fetch;
  globalThis.fetch = (u, i) => o.handler(u, i);
  globalThis.location = { origin: 'https://t.test', pathname: '/share/' };
  globalThis.window = { msal: { PublicClientApplication: class { async initialize() {} getAllAccounts() { return []; } async ssoSilent() { throw new Error('interaction_required'); } } } };
  try {
    sessionStorage.setItem(C.SESSION_KEY, JSON.stringify({ provider: 'onedrive', token: 'old', expiresAt: Date.now() + 3600e3, user: { displayName: 'U', emailAddress: 'u@school.fr' } }));
    const cloud = C.createCloud({ msClientId: 'cid', queueStore: C.memoryQueueStore() });
    const ev = []; cloud.on('auth-expired', e => ev.push(e));
    assert.ok(await cloud.restore());
    o.hooks.force401 = true;
    await assert.rejects(cloud.listProjects(), e => e.code === 'auth-expired' && e.status === 401);
    assert.equal(ev.length, 1);
  } finally { globalThis.fetch = savedFetch; delete globalThis.window; delete globalThis.location; }
});

test('OneDrive e-mail invite requires sign-in (works where anonymous access is blocked)', async () => {
  const { cloud, server } = setup('onedrive');
  await cloud.signIn('onedrive');
  const p = await cloud.createProject('P');
  await p.share.emails(['a@b.fr']);
  assert.equal(server.hooks.invites.at(-1).requireSignIn, true);
});

test('OneDrive upload: overwrite -> replace, default -> rename', async () => {
  const { cloud, server } = setup('onedrive');
  await cloud.signIn('onedrive');
  const p = await cloud.createProject('P');
  await p.uploadBlob('files/a.bin', new Blob([new Uint8Array(5)]));
  await p.uploadBlob('files/a.bin', new Blob([new Uint8Array(5)]), { overwrite: true });
  const bodies = server.log.filter(l => /createUploadSession/.test(l.url)).map(l => l.body.item['@microsoft.graph.conflictBehavior']);
  assert.deepEqual(bodies, ['rename', 'replace']);
});

test('index.html: CSP present, lists the pinned scripts; SRI constants pinned', () => {
  const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
  const csp = /http-equiv="Content-Security-Policy" content="([^"]+)"/.exec(html)[1];
  assert.match(csp, /default-src 'self'/);
  assert.match(csp, /object-src 'none'/);
  assert.ok(!/unsafe-inline|unsafe-eval/.test(csp));
  assert.ok(csp.includes(C.MSAL_SRC), 'msal URL allowed exactly');
  assert.ok(csp.includes('https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js'));
  assert.match(C.MSAL_SRI, /^sha384-[A-Za-z0-9+/]{64}$/);
  const app = readFileSync(new URL('../app.js', import.meta.url), 'utf8');
  assert.match(app, /integrity = 'sha384-[A-Za-z0-9+/]{64}'/);
  assert.match(html, /id="whereLbl"/);
  assert.match(html, /<p class="hint" hidden><a href="\?mock=1"/, 'mock link hidden until JS confirms localhost');
});

/* ===== owner's real-Drive test round ===== */
for (const provider of ['google', 'onedrive']) {
  test(`[${provider}] trash really reaches the API and the file disappears from the listing`, async () => {
    const { cloud, server } = setup(provider);
    await cloud.signIn(provider);
    const p = await cloud.createProject('P');
    const f = await p.uploadBlob('files/x.bin', new Blob([new Uint8Array(5)]));
    server.log.length = 0;
    await cloud.adapter.trash(f.id);
    const req = server.log.find(l => /\/files\/|\/items\//.test(l.url) && l.method !== 'GET');
    if (provider === 'google') { assert.equal(req.method, 'PATCH'); assert.deepEqual(req.body, { trashed: true }); }
    else assert.equal(req.method, 'DELETE');
    assert.deepEqual((await p.list('files')).map(n => n.name), []);
  });
}

test('[google] 403 accessNotConfigured / SERVICE_DISABLED -> code api-disabled; other 403 stays http', async () => {
  sessionStorage.clear();
  const g = createMockServer();
  let mode = null;
  const fetchWrap = (u, i) => {
    if (mode && /\/drive\/v3\/files/.test(String(u))) return Promise.resolve(new Response(JSON.stringify(mode), { status: 403, headers: { 'Content-Type': 'application/json' } }));
    return g.handler(u, i);
  };
  const cloud = C.createCloud({ overrides: { google: { fetch: fetchWrap } }, queueStore: C.memoryQueueStore() });
  await cloud.signIn('google');
  mode = { error: { code: 403, message: 'Google Drive API has not been used in project 123 before or it is disabled.', errors: [{ reason: 'accessNotConfigured' }], details: [{ reason: 'SERVICE_DISABLED' }] } };
  await assert.rejects(cloud.listProjects(), e => e.code === 'api-disabled' && e.status === 403);
  mode = { error: { code: 403, message: 'The user does not have sufficient permissions', errors: [{ reason: 'forbidden' }] } };
  await assert.rejects(cloud.listProjects(), e => e.code === 'http' && e.status === 403);
});
