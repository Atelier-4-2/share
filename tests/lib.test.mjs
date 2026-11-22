import test from 'node:test';
import assert from 'node:assert/strict';
import * as L from '../lib.js';

test('chunkCount / nextRange', () => {
  assert.equal(L.chunkCount(0), 1);
  assert.equal(L.chunkCount(1), 1);
  assert.equal(L.chunkCount(L.CHUNK_SIZE), 1);
  assert.equal(L.chunkCount(L.CHUNK_SIZE + 1), 2);
  assert.equal(L.chunkCount(500 * 1024 * 1024), 63);
  const r = L.nextRange(0, 20, 8);
  assert.deepEqual([r.start, r.end, r.contentRange], [0, 8, 'bytes 0-7/20']);
  const last = L.nextRange(16, 20, 8);
  assert.equal(last.contentRange, 'bytes 16-19/20');
  assert.equal(L.nextRange(0, 0).contentRange, 'bytes */0');
  assert.equal(L.CHUNK_SIZE % (256 * 1024), 0);
});
test('Range header → offset', () => {
  assert.equal(L.parseRangeHeader('bytes=0-8388607'), 8388608);
  assert.equal(L.parseRangeHeader(null), 0);
  assert.equal(L.statusRangeHeader(99), 'bytes */99');
});
test('resume key', () => {
  const f = { name: 'a.3dm', size: 10, lastModified: 5 };
  assert.equal(L.resumeKey(f, 'P1', 'x/a.3dm'), 'share-resume|P1|x/a.3dm|10|5');
  assert.notEqual(L.resumeKey(f, 'P1', 'a'), L.resumeKey({ ...f, size: 11 }, 'P1', 'a'));
  assert.notEqual(L.resumeKey(f, 'P1', 'a'), L.resumeKey({ ...f, lastModified: 6 }, 'P1', 'a'));
});
test('backoff and retryable', () => {
  assert.equal(L.backoffDelay(1, 0.5), 800);
  assert.equal(L.backoffDelay(2, 0.5), 1600);
  assert.ok(L.backoffDelay(20, 1) <= 30000 * 1.25);
  assert.ok(L.isRetryable(503) && L.isRetryable(429) && L.isRetryable(0));
  assert.ok(!L.isRetryable(404) && !L.isRetryable(403));
});
test('emails', () => {
  assert.ok(L.isValidEmail('a@b.fr'));
  assert.ok(!L.isValidEmail('a@b'));
  assert.ok(!L.isValidEmail('a b@c.com'));
  const r = L.parseEmails('a@b.fr, C@d.com\nA@B.fr; bad, <e@f.kr>');
  assert.deepEqual(r.valid, ['a@b.fr', 'C@d.com', 'e@f.kr']);
  assert.deepEqual(r.invalid, ['bad']);
  assert.deepEqual(r.duplicates, ['A@B.fr']);
  assert.deepEqual(L.parseEmails('').valid, []);
});
test('folder paths parents first', () => {
  const p = L.folderPaths([{ path: 'a/b/c/x.txt' }, { path: 'a/y.txt' }, { path: 'z.txt' }, { path: 'd/e/f.txt' }]);
  assert.deepEqual(p, ['a', 'd', 'a/b', 'd/e', 'a/b/c']);
  assert.deepEqual(L.splitPath('a/b/c.txt'), { dir: 'a/b', name: 'c.txt' });
  assert.deepEqual(L.splitPath('c.txt'), { dir: '', name: 'c.txt' });
});
test('file tree', () => {
  const F = L.FOLDER_MIME;
  const items = [
    { id: '2', name: 'sub', mimeType: F, parents: ['1'] },
    { id: '3', name: 'b.txt', mimeType: 'text/plain', parents: ['2'], size: '10' },
    { id: '4', name: 'a.txt', mimeType: 'text/plain', parents: ['1'], size: '5' },
  ];
  const tree = L.buildFileTree('1', items);
  assert.equal(tree[0].name, 'sub'); // folders first
  assert.equal(tree[0].children[0].name, 'b.txt');
  assert.equal(L.treeSize(tree), 15);
  assert.equal(L.treeCount(tree), 2);
});
test('escapeQuery / safeName', () => {
  assert.equal(L.escapeQuery(String.raw`it's \ x`), String.raw`it\'s \\ x`);
  assert.equal(L.safeName(' a/b '), 'a-b');
});
test('formatters', () => {
  assert.equal(L.formatBytes(0), '0 B');
  assert.equal(L.formatBytes(1536), '1.50 KB');
  assert.equal(L.formatBytes(500 * 1024 * 1024), '500 MB');
  assert.equal(L.formatEta(75), '1m 15s');
  assert.equal(L.formatEta(Infinity), '–');
  assert.equal(L.speedFromSamples([{ t: 0, b: 0 }, { t: 2000, b: 2000 }]), 1000);
});
test('OneDrive 320 KiB alignment', () => {
  assert.equal(L.ONEDRIVE_CHUNK, 10 * 1024 * 1024);
  assert.ok(L.isGraphAligned(L.ONEDRIVE_CHUNK));
  assert.ok(!L.isGraphAligned(1000000));
  assert.equal(L.alignDown(1000000), 983040);
  assert.equal(L.alignDown(10), 327680);
  assert.equal(L.parseNextExpected(['26-']), 26);
  assert.equal(L.parseNextExpected(['327680-655359', '900-']), 327680);
  assert.equal(L.parseNextExpected([]), 0);
});
test('admin consent detection', () => {
  assert.ok(L.isAdminConsentError('AADSTS90094: The grant requires admin permission'));
  assert.ok(L.isAdminConsentError('AADSTS65001: The user has not consented'));
  assert.ok(L.isAdminConsentError('consent_required'));
  assert.ok(!L.isAdminConsentError('network error'));
});
test('graph permission mapping', () => {
  assert.deepEqual(L.mapGraphPermission({ id: 'a', roles: ['read'], link: { scope: 'anonymous', webUrl: 'u' } }), { id: 'a', type: 'anyone', role: 'reader', url: 'u' });
  assert.equal(L.mapGraphPermission({ id: 'b', roles: ['read'], link: { scope: 'organization', webUrl: 'u' } }).type, 'organization');
  const u = L.mapGraphPermission({ id: 'c', roles: ['write'], invitation: { email: 'x@y.fr' } });
  assert.deepEqual([u.type, u.role, u.emailAddress], ['user', 'writer', 'x@y.fr']);
  assert.equal(L.mapGraphPermission({ id: 'd', roles: ['owner'], grantedToV2: { user: { email: 'o@p.q', displayName: 'O' } } }).role, 'owner');
});
