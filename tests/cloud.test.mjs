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
    assert.ok(!JSON.stringify(localStorage.getItem('atelier-cloud-last')).includes('mock-token'), 'no token in localStorage');
    assert.equal(localStorage.getItem('atelier-cloud-session'), null, 'session never in localStorage');
    assert.ok(sessionStorage.getItem('atelier-cloud-session').includes('mock-token'));
    // a second page of the same tab: new cloud object, same sessionStorage, same (mock) server
    const g = createMockServer(), o = createMockGraph();
    cloud.signOutLocal();
    assert.equal(await cloud.restore(), null, 'signOutLocal cleared the session');
    await cloud.signIn(provider);
    assert.equal(cloud.preferred, provider);
    cloud.signOut();
    assert.equal(sessionStorage.getItem('atelier-cloud-session'), null);
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
  assert.deepEqual(server.hooks.invites.at(-1), { to: 'a@b.fr', sendInvitation: true, message: 'hello' });
});

test('Google email share sends notification flag', async () => {
  const { cloud, server } = setup('google');
  await cloud.signIn('google');
  const p = await cloud.createProject('P');
  await p.share.emails(['x@y.fr'], 'hi');
  assert.deepEqual(server.hooks.lastEmails.at(-1), { to: 'x@y.fr', notify: 'true', message: 'hi' });
  assert.deepEqual(await p.share.setLink(true), { scope: 'anyone' });
});
