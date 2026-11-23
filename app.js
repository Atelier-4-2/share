import {
  parseEmails, folderPaths, splitPath, safeName, safePath, isLocalHost, safeHttpsUrl, isAuthPopup, formatBytes, formatSpeed, formatEta, treeSize, treeCount,
} from './lib.js';
import { createCloud, CloudError, MOCK_SESSION_KEY } from './atelier-cloud.js';
import { UploadTask, UploadQueue } from './upload.js';
import { DICT } from './i18n.js';

let lang = 'ko';
try { const s = localStorage.getItem('atelier-lang'); if (s === 'fr' || s === 'ko') lang = s; } catch (e) { /* storage unavailable */ }
const t = (k, p) => {
  let s = (DICT[lang] && DICT[lang][k]) ?? DICT.ko[k] ?? k;
  if (p) for (const key in p) s = s.split('{' + key + '}').join(p[key]);
  return s;
};
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c != null) el.append(c.nodeType ? c : document.createTextNode(String(c)));
  return el;
}

/* ---------------- state ---------------- */
const cfg = window.ATELIER_SHARE_CONFIG || {};
const qs = new URLSearchParams(location.search);
// the in-memory fake storage exists for development only: ?mock=1 is ignored on any real host
const LOCAL = isLocalHost(location.hostname);
const MOCK = LOCAL && qs.get('mock') === '1';
void MOCK_SESSION_KEY;
const ROOT_NAME = cfg.ROOT_FOLDER_NAME || 'Atelier 4·2 Share';
const state = { adapter: null, user: null, quota: null, rootId: null, projects: [], project: null, tree: [], perms: [], tasks: [], folderCache: new Map(), linkScope: null, provider: null };
const queue = new UploadQueue(2);
const store = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
};
const PROVIDER_KEY = 'share-provider';

let toastTimer;
function toast(msg) {
  const el = $('#toast'); el.textContent = msg; el.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => el.classList.remove('show'), 3600);
}
const errMsg = (e) => (e && (e.message || e.status)) || 'error';
const fmtDate = (s) => { try { return new Date(s).toLocaleDateString(lang === 'fr' ? 'fr-FR' : 'ko-KR', { year: 'numeric', month: 'short', day: 'numeric' }); } catch (e) { return ''; } };
const pname = () => (state.adapter ? state.adapter.label : '');

function configured(provider) { return MOCK || !!(provider === 'google' ? cfg.GOOGLE_CLIENT_ID : cfg.MS_CLIENT_ID); }

function applyLang() {
  document.documentElement.lang = lang;
  $$('[data-i18n]').forEach(el => { el.textContent = t(el.dataset.i18n); });
  $$('[data-i18n-aria]').forEach(el => el.setAttribute('aria-label', t(el.dataset.i18nAria)));
  $$('[data-i18n-ph]').forEach(el => el.setAttribute('placeholder', t(el.dataset.i18nPh)));
  $('#langFr').setAttribute('aria-pressed', String(lang === 'fr'));
  $('#langKo').setAttribute('aria-pressed', String(lang === 'ko'));
  document.title = 'Share — Atelier 4·2';
  if (MOCK) { const b = $('#mockBanner'); b.hidden = false; b.textContent = t('mockBanner'); }
  $('#tryMock').parentElement.hidden = !LOCAL || MOCK; // never shown on the published site
  renderChooser();
  if (state.user) { renderAccount(); renderProjects(); renderProviderTexts(); }
  if (state.project) { renderProjectHead(); renderUploads(); renderPerms(); renderTree(); }
  if (lastQueue) renderQueue(lastQueue);
}
$$('.lang button').forEach(b => b.addEventListener('click', () => {
  lang = b.dataset.lang;
  try { localStorage.setItem('atelier-lang', lang); } catch (e) { /* ignore */ }
  applyLang();
}));

/* ---------------- chooser / sign-in ---------------- */
function show(which) {
  $('#screenSignin').hidden = which !== 'signin';
  $('#screenApp').hidden = which !== 'app';
}
function renderChooser() {
  const last = store.get(PROVIDER_KEY);
  let anyMissing = false;
  for (const p of ['google', 'onedrive']) {
    const b = $(p === 'google' ? '#provGoogle' : '#provOnedrive');
    const ok = configured(p);
    if (!ok) anyMissing = true;
    b.disabled = !ok;
    b.setAttribute('aria-pressed', String(last === p));
    $('[data-state]', b).textContent = !ok ? t('needSetup') : last === p ? t('lastUsed') : '';
  }
  $('#setupNote').hidden = !anyMissing || MOCK;
}

let cloud = null, watcher = null;
const mockServers = {};
async function initCloud() {
  const overrides = {};
  if (MOCK) {
    const { createMockServer } = await import('./mock-server.js');
    const { createMockGraph } = await import('./mock-graph.js');
    mockServers.google = createMockServer(); mockServers.onedrive = createMockGraph();
    window.__mockServers = mockServers; // dev hooks, localhost only
    overrides.google = { fetch: (u, i) => mockServers.google.handler(u, i), refuseRefresh: () => !!window.__refuseRefresh };
    overrides.onedrive = {
      fetch: (u, i) => mockServers.onedrive.handler(u, i), refuseRefresh: () => !!window.__refuseRefresh,
      signIn: async () => { if (qs.get('mockadmin') === '1') throw new CloudError('admin-consent', 'AADSTS90094: Need admin approval (mock)'); },
    };
  }
  cloud = createCloud({
    googleClientId: cfg.GOOGLE_CLIENT_ID, msClientId: cfg.MS_CLIENT_ID, rootName: ROOT_NAME, overrides, mock: MOCK,
    isOffline: MOCK ? () => !!window.__forceOffline || navigator.onLine === false : undefined,
    pollMs: MOCK && qs.get('poll') ? Number(qs.get('poll')) : undefined,
  });
  if (MOCK) window.__cloud = cloud;
  cloud.on('auth-expired', () => { $('#reauthBar').hidden = false; });
  cloud.on('auth-restored', () => {
    $('#reauthBar').hidden = true; toast(t('reauthed'));
    // uploads that stopped for the expired sign-in carry on from their stored session
    state.tasks.forEach(x => { if (x.state === 'error' && x.needsAuth) x.resume(); });
    queue.pump();
  });
  cloud.on('queue', renderQueue);
  cloud.on('queue', (s) => { if (s.pending === 0 && state.project) loadTree(); });
}
$('#reauthBtn').addEventListener('click', async () => { try { await cloud.reauth(); } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); } });

async function enterApp(user) {
  state.adapter = cloud.adapter; state.provider = cloud.provider; state.user = user;
  state.quota = await cloud.quota();
  if (MOCK) { const m = mockServers[state.provider]; window.__mockLog = m.log; window.__mock = m; }
  show('app'); renderAccount(); renderProviderTexts();
  $('#mockBig').hidden = !MOCK;
  renderChooser();
  await loadProjects();
  renderQueue(await cloud.queueStatus());
  // pick up where this device left off
  const last = cloud.lastProject();
  if (last && last.provider === cloud.provider) { const p = state.projects.find(x => x.id === last.projectId); if (p) openProject(p); }
}

async function chooseProvider(provider) {
  const errEl = $('#signInError'); errEl.hidden = true;
  $$('.prov').forEach(b => { b.disabled = true; });
  try {
    store.set(PROVIDER_KEY, provider);
    await enterApp(await cloud.signIn(provider));
  } catch (e) {
    const net = e.code === 'network' ? (provider === 'google' ? t('gsiFail') : t('msalFail')) : null;
    errEl.textContent = e.code === 'admin-consent' ? t('adminConsent') : net || (t('signInFail') + ' (' + errMsg(e) + ')');
    errEl.hidden = false; state.adapter = null; state.user = null; show('signin');
  } finally { renderChooser(); }
}
$('#provGoogle').addEventListener('click', () => chooseProvider('google'));
$('#provOnedrive').addEventListener('click', () => chooseProvider('onedrive'));

function stopWatch() { if (watcher) { watcher.stop(); watcher = null; } }
function leaveStore() {
  stopWatch();
  state.tasks.forEach(x => { if (['uploading', 'retrying', 'queued'].includes(x.state)) x.pause(); });
  if (cloud) cloud.signOut();
  Object.assign(state, { adapter: null, user: null, quota: null, rootId: null, projects: [], project: null, tree: [], perms: [], tasks: [], folderCache: new Map(), provider: null });
  queue.tasks = [];
  $('#reauthBar').hidden = true; $('#queueBar').hidden = true;
  $('#viewProject').hidden = true; $('#viewProjects').hidden = false;
  show('signin'); renderChooser();
}
$('#signOutBtn').addEventListener('click', leaveStore);
$('#switchBtn').addEventListener('click', leaveStore);

let lastQueue = null;
function renderQueue(s) {
  lastQueue = s;
  const bar = $('#queueBar');
  if (!s || (!s.pending && !s.conflicts && !s.failed)) { bar.hidden = true; bar.replaceChildren(); return; }
  const parts = [];
  if (s.pending) parts.push(t('queueBar', { n: s.pending }));
  if (s.conflicts) parts.push(t('queueConflict', { n: s.conflicts }));
  if (s.failed) parts.push(t('queueFailed', { n: s.failed }));
  const list = h('ul', { class: 'queue-items' });
  const act = async (id, how) => {
    try { await cloud.resolveQueued(id, how); } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); }
  };
  for (const it of (s.items || []).filter(x => x.state === 'conflict' || x.state === 'failed')) {
    const btns = h('span', { class: 'acts' });
    if (it.state === 'conflict') btns.append(
      h('button', { type: 'button', class: 'btn sm', 'data-q': 'overwrite', onclick: () => act(it.id, 'overwrite') }, t('qOverwrite')),
      ...(it.kind === 'json' ? [h('button', { type: 'button', class: 'btn sm', 'data-q': 'keep-both', onclick: () => act(it.id, 'keep-both') }, t('qKeepBoth'))] : []));
    else btns.append(h('button', { type: 'button', class: 'btn sm', 'data-q': 'retry', onclick: () => act(it.id, 'retry') }, t('retry')));
    btns.append(h('button', { type: 'button', class: 'btn sm danger', 'data-q': 'discard', onclick: () => act(it.id, 'discard') }, t('qDiscard')));
    list.append(h('li', { 'data-state': it.state }, h('span', { class: 'q-path' }, (it.projectName ? it.projectName + ' / ' : '') + it.path, ' · ', it.state === 'conflict' ? t('qStateConflict') : t('qStateFailed')), btns));
  }
  bar.replaceChildren(h('span', {}, parts.join(' ')), list.children.length ? list : '');
  bar.hidden = false;
}

function renderAccount() {
  if (!state.user) return;
  $('#userName').textContent = state.user.displayName || '';
  $('#userMail').textContent = state.user.emailAddress || '';
  $('#provBadge').textContent = pname();
  const q = state.quota;
  $('#quota').textContent = !q ? '' : q.limit ? t('quotaOf', { used: formatBytes(q.usage), limit: formatBytes(q.limit), name: pname() }) : t('quota', { used: formatBytes(q.usage), name: pname() });
}
function renderProviderTexts() {
  const a = state.adapter; if (!a) return;
  $('#projectWhere').textContent = t('projectWhere', { name: a.label, root: a.id === 'google' ? ROOT_NAME : t('appFolder') });
  $('#onlyApp').textContent = t('onlyApp');
  $('#mailNote').textContent = t('mailNote', { by: a.mailBy });
  $('#openDrive').textContent = t('openIn', { name: a.label });
}

/* ---------------- projects ---------------- */
async function loadProjects() {
  try { state.projects = await state.adapter.listProjects(); }
  catch (e) { toast(t('loadErr', { msg: errMsg(e) })); state.projects = []; }
  renderProjects();
}
function renderProjects() {
  const ul = $('#projectList'); ul.replaceChildren();
  for (const p of state.projects) {
    ul.append(h('li', {}, h('button', { type: 'button', class: 'open', onclick: () => openProject(p) }, p.name), h('span', { class: 'when' }, fmtDate(p.createdTime))));
  }
  $('#projectsEmpty').hidden = state.projects.length > 0;
}
$('#newProjectForm').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const err = $('#projectError'); err.hidden = true;
  const name = safeName($('#projectName').value);
  if (!name) return;
  try {
    const p = await cloud.createProject(name);
    $('#projectName').value = '';
    state.projects.push(p.node); state.projects.sort((a, b) => a.name.localeCompare(b.name));
    renderProjects(); toast(t('projectCreated'));
    openProject(p);
  } catch (e) {
    err.textContent = e.code === 'exists' ? t('projectExists') : t('apiErr', { msg: errMsg(e) }); err.hidden = false;
  }
});

async function openProject(p) {
  stopWatch();
  let handle;
  try { handle = p.folderId ? p : await cloud.openProject(p); } catch (e) { return toast(t('loadErr', { msg: errMsg(e) })); }
  state.project = handle; state.tree = []; state.perms = []; state.tasks = []; queue.tasks = [];
  state.filesId = await handle.folderId('files'); state.folderCache = new Map([['', state.filesId]]);
  $('#viewProjects').hidden = true; $('#viewProject').hidden = false;
  renderProjectHead(); renderUploads();
  $('#linkBox').hidden = true; $('#linkOff').hidden = false; $('#shareResult').hidden = true;
  window.scrollTo(0, 0);
  cloud.rememberProject(handle, 'share');
  await Promise.all([loadTree(), loadPerms()]);
  // other devices: refresh the list when something changes (20 s poll while visible, immediately on tab focus)
  watcher = handle.watch(() => {
    const quiet = Date.now() - (state.ownAt || 0) < 30000;
    loadTree(); loadPerms();
    if (!quiet) toast(t('remoteChanged'));
  });
}
$('#backBtn').addEventListener('click', () => {
  stopWatch();
  state.tasks.forEach(x => { if (['uploading', 'retrying', 'queued'].includes(x.state)) x.pause(); });
  state.project = null; $('#viewProject').hidden = true; $('#viewProjects').hidden = false; loadProjects();
});
function projectUrl() { return state.project.openUrl(); }
function renderProjectHead() {
  $('#projTitle').textContent = state.project.name;
  const u = safeHttpsUrl(projectUrl());
  if (u) $('#openDrive').setAttribute('href', u); else $('#openDrive').removeAttribute('href');
}

/* ---------------- tree ---------------- */
async function loadTree() {
  try { state.tree = await state.project.tree('files'); }
  catch (e) { toast(t('loadErr', { msg: errMsg(e) })); }
  renderTree();
}
$('#refreshBtn').addEventListener('click', loadTree);

function renderTree() {
  const root = $('#tree'); root.replaceChildren();
  const n = treeCount(state.tree);
  $('#treeEmpty').hidden = state.tree.length > 0;
  $('#totals').textContent = t('totals', { n, size: formatBytes(treeSize(state.tree)) });
  if (!state.tree.length) return;
  root.append(renderNodes(state.tree));
}
function renderNodes(nodes) {
  const ul = h('ul');
  for (const n of nodes) {
    const li = h('li');
    li.append(nodeRow(n));
    if (n.isFolder && n.children && n.children.length) li.append(renderNodes(n.children));
    ul.append(li);
  }
  return ul;
}
function nodeRow(n) {
  const row = h('div', { class: 'node ' + (n.isFolder ? 'folder' : 'file'), 'data-id': n.id });
  const href = n.isFolder ? '' : safeHttpsUrl(state.adapter.openUrl(n));
  const nameEl = h('span', { class: 'n-name' }, n.isFolder ? h('span', {}, n.name + '/') : href ? h('a', { href, target: '_blank', rel: 'noopener', title: t('openIn', { name: pname() }) }, n.name) : h('span', {}, n.name));
  const meta = h('span', { class: 'n-meta' }, n.isFolder ? '' : formatBytes(n.size) + ' · ' + fmtDate(n.createdTime));
  const acts = h('span', { class: 'acts' });
  const renameBtn = h('button', { type: 'button', class: 'btn sm', onclick: () => startRename(row, n, nameEl) }, t('rename'));
  const delBtn = h('button', { type: 'button', class: 'btn sm danger', onclick: () => confirmDelete(delBtn, n) }, t('del'));
  acts.append(renameBtn, delBtn);
  row.append(nameEl, meta, acts);
  return row;
}
function startRename(row, n, nameEl) {
  const input = h('input', { type: 'text', class: 'rename', value: n.name, 'aria-label': t('rename') });
  const save = h('button', { type: 'button', class: 'btn sm primary' }, t('save'));
  const doSave = async () => {
    const name = safeName(input.value);
    if (!name || name === n.name) return loadTree();
    try { state.ownAt = Date.now(); await state.adapter.rename(n.id, name); toast(t('renamed')); } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); }
    loadTree();
  };
  save.addEventListener('click', doSave);
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') doSave(); if (e.key === 'Escape') loadTree(); });
  nameEl.replaceChildren(input); row.querySelector('.acts').replaceChildren(save); input.focus(); input.select();
}
function confirmDelete(btn, n) {
  if (btn.dataset.armed) {
    state.adapter.trash(n.id).then(() => { toast(t('deleted')); return loadTree(); }).catch(e => toast(t('apiErr', { msg: errMsg(e) })));
    return;
  }
  btn.dataset.armed = '1'; btn.textContent = t('confirmDel');
  setTimeout(() => { if (btn.isConnected) { delete btn.dataset.armed; btn.textContent = t('del'); } }, 3500);
}

/* ---------------- file intake (drop / pick) ---------------- */
const drop = $('#drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'dragleave' && drop.contains(e.relatedTarget)) return; drop.classList.remove('over'); }));
drop.addEventListener('drop', async (e) => {
  const items = e.dataTransfer && e.dataTransfer.items;
  let entries;
  if (items && items.length && items[0].webkitGetAsEntry) {
    // entries must be grabbed synchronously, before any await
    const roots = [...items].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
    entries = await walkEntries(roots);
  } else {
    entries = [...(e.dataTransfer.files || [])].map(f => ({ file: f, path: f.name }));
  }
  addFiles(entries);
});
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('#fileInput').click(); } });
$('#pickFiles').addEventListener('click', () => $('#fileInput').click());
$('#pickFolder').addEventListener('click', () => $('#folderInput').click());
$('#fileInput').addEventListener('change', (e) => { addFiles([...e.target.files].map(f => ({ file: f, path: f.name }))); e.target.value = ''; });
$('#folderInput').addEventListener('change', (e) => { addFiles([...e.target.files].map(f => ({ file: f, path: f.webkitRelativePath || f.name }))); e.target.value = ''; });
$('#mockBig').addEventListener('click', () => {
  const mb = new Uint8Array(1024 * 1024);
  const file = new File(new Array(500).fill(mb), 'render_4k_fake_500MB.mov', { type: 'video/quicktime', lastModified: 1700000000000 });
  addFiles([{ file, path: 'renders/' + file.name }]);
});
window.__share = { addFiles: (e) => addFiles(e) }; // test hook (also used for synthetic folder drops)

/** Recursively read FileSystemEntry objects (readEntries returns batches until empty). */
async function walkEntries(roots) {
  const out = [];
  const readAll = (reader) => new Promise((res, rej) => {
    const acc = [];
    const next = () => reader.readEntries((batch) => { if (!batch.length) res(acc); else { acc.push(...batch); next(); } }, rej);
    next();
  });
  const visit = async (entry, prefix) => {
    if (entry.isFile) {
      const file = await new Promise((res, rej) => entry.file(res, rej));
      out.push({ file, path: prefix + entry.name });
    } else if (entry.isDirectory) {
      const kids = await readAll(entry.createReader());
      if (!kids.length) out.push({ dir: true, path: prefix + entry.name + '/' });
      for (const k of kids) await visit(k, prefix + entry.name + '/');
    }
  };
  for (const r of roots) await visit(r, '');
  return out;
}

async function ensurePath(dir) {
  const cache = state.folderCache;
  if (cache.has(dir)) return cache.get(dir);
  const { dir: parent, name } = splitPath(dir);
  const parentId = await ensurePath(parent);
  const f = await state.adapter.ensureFolder(name, parentId);
  cache.set(dir, f.id); return f.id;
}

async function addFiles(entries) {
  if (!state.project) return;
  entries = entries.filter(e => e.path);
  if (!entries.length) return toast(t('nothingToUpload'));
  const project = state.project;
  if (cloud.offline) {
    // no network: keep the files in IndexedDB; the cloud layer sends them when the connection is back
    let n = 0;
    for (const e of entries) { if (e.dir) continue; const sp = safePath(e.path); if (!sp) continue; await project.saveBlob('files/' + sp, e.file); n++; }
    toast(t('queuedOffline', { n })); renderQueue(await cloud.queueStatus());
    return;
  }
  toast(t('preparing'));
  try {
    // stored names are cleaned (OneDrive rejects : * ? " < > | and trailing dots); the list keeps showing the original path
    entries = entries.map(e => ({ ...e, orig: e.path, path: e.dir ? safePath(e.path) + '/' : safePath(e.path) })).filter(e => e.path && e.path !== '/');
    const dirs = folderPaths(entries.map(e => ({ path: e.dir ? e.path + 'x' : e.path })));
    for (const d of dirs) await ensurePath(d);
    let n = 0;
    for (const e of entries) {
      if (e.dir) continue;
      const { dir } = splitPath(e.path);
      const parentId = await ensurePath(dir);
      const task = new UploadTask({ drive: state.adapter, file: e.file, parentId, relPath: e.path, label: e.orig, onUpdate: onTaskUpdate });
      state.tasks.push(task); queue.add(task); n++;
    }
    if (state.project === project) { renderUploads(); toast(t('added', { n })); }
  } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); }
}

/* ---------------- upload list ---------------- */
let rafPending = false;
function onTaskUpdate(task) {
  state.ownAt = Date.now();
  if (['done', 'error', 'paused', 'canceled'].includes(task.state)) queue.pump();
  if (task.state === 'done' && state.tasks.every(x => ['done', 'canceled', 'error', 'paused'].includes(x.state))) { loadTree(); refreshQuota(); }
  if (rafPending) return; rafPending = true;
  setTimeout(() => { rafPending = false; renderUploads(); }, 120);
}
async function refreshQuota() { try { state.quota = await state.adapter.quota(); renderAccount(); } catch (e) { /* ignore */ } }

function statusText(task) {
  switch (task.state) {
    case 'queued': return t('sQueued');
    case 'uploading': return (task.resumedFromStore && task.sent > 0 ? t('resumedNote') + ' · ' : '') + t('sUploading') + ' · ' + formatSpeed(task.speed) + ' · ' + formatEta(task.eta);
    case 'retrying': return t('sRetrying', { n: task.attempt, max: 6 });
    case 'paused': return t('sPaused');
    case 'done': return t('sDone');
    case 'error': return task.needsAuth ? t('sNeedAuth') : t('sError') + (task.error ? ' (' + errMsg(task.error) + ')' : '');
    case 'canceled': return t('sCanceled');
    default: return '';
  }
}
function renderUploads() {
  const ul = $('#upList');
  const live = new Map([...ul.children].map(li => [li.dataset.id, li]));
  for (const task of state.tasks) {
    let li = live.get(String(task.id));
    if (!li) {
      li = h('li', { class: 'up-item', 'data-id': task.id },
        h('div', { class: 'up-top' }, h('span', { class: 'up-name' }), h('span', { class: 'up-meta pct' })),
        h('div', { class: 'bar', role: 'progressbar', 'aria-valuemin': '0', 'aria-valuemax': '100' }, h('i')),
        h('div', { class: 'up-meta status' }),
        h('div', { class: 'up-actions' }));
      ul.append(li);
    }
    live.delete(String(task.id));
    const pct = task.size ? Math.min(100, Math.floor(task.sent / task.size * 100)) : (task.state === 'done' ? 100 : 0);
    li.className = 'up-item ' + task.state;
    $('.up-name', li).textContent = task.label || task.relPath;
    $('.pct', li).textContent = formatBytes(task.sent) + ' / ' + formatBytes(task.size) + ' · ' + pct + '%';
    $('.bar > i', li).style.width = pct + '%';
    $('.bar', li).setAttribute('aria-valuenow', String(pct));
    $('.bar', li).setAttribute('aria-label', task.label || task.relPath);
    $('.status', li).textContent = statusText(task);
    const acts = $('.up-actions', li); acts.replaceChildren();
    if (['uploading', 'retrying', 'queued'].includes(task.state)) acts.append(h('button', { type: 'button', class: 'btn sm', 'data-act': 'pause', onclick: () => task.pause() }, t('pause')));
    if (task.state === 'paused') acts.append(h('button', { type: 'button', class: 'btn sm', 'data-act': 'resume', onclick: () => { task.resume(); queue.pump(); } }, t('resume')));
    if (task.state === 'error' && task.needsAuth) acts.append(h('button', { type: 'button', class: 'btn sm primary', 'data-act': 'reauth', onclick: async () => { try { await cloud.reauth(); } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); } } }, t('reauthBtn')));
    else if (task.state === 'error') acts.append(h('button', { type: 'button', class: 'btn sm', 'data-act': 'retry', onclick: () => task.resume() }, t('retry')));
    if (!['done', 'canceled'].includes(task.state)) acts.append(h('button', { type: 'button', class: 'btn sm', 'data-act': 'cancel', onclick: () => task.cancel() }, t('cancel')));
  }
  for (const li of live.values()) li.remove();
}

/* ---------------- sharing ---------------- */
async function loadPerms() {
  try { state.perms = await state.project.share.list(); } catch (e) { toast(t('loadErr', { msg: errMsg(e) })); state.perms = []; }
  renderPerms();
}
const roleLabel = (r) => r === 'owner' ? t('roleOwner') : r === 'writer' ? t('roleWriter') : t('roleReader');
function renderPerms() {
  const ul = $('#permList'); ul.replaceChildren();
  const link = state.perms.find(p => p.type === 'anyone') || state.perms.find(p => p.type === 'organization');
  $('#linkBox').hidden = !link; $('#linkOff').hidden = !!link;
  $('#orgNote').hidden = !(link && link.type === 'organization');
  if (link) {
    const url = safeHttpsUrl(link.url);   // only https links from the provider are shown, copied, mailed or turned into a QR
    $('#linkInput').value = url;
    if (url) $('#mailtoLink').setAttribute('href', 'mailto:?subject=' + encodeURIComponent(t('mailSubject', { name: state.project.name })) + '&body=' + encodeURIComponent(t('mailBody', { name: state.project.name, url })));
    else $('#mailtoLink').removeAttribute('href');
    drawQr(url);
  }
  for (const p of state.perms) {
    const label = p.type === 'anyone' ? t('anyoneLabel') : p.type === 'organization' ? t('orgLabel') : (p.emailAddress || p.displayName || p.id);
    const li = h('li', { 'data-perm': p.id },
      h('div', { class: 'p-main' }, h('div', {}, label), h('div', { class: 'p-role' }, roleLabel(p.role))));
    if (p.role !== 'owner') li.append(h('button', { type: 'button', class: 'btn sm', onclick: () => removePerm(p) }, t('remove')));
    ul.append(li);
  }
  $('#permsEmpty').hidden = state.perms.filter(p => p.role !== 'owner').length > 0;
}
async function removePerm(p) {
  try { await state.project.share.remove(p.id); toast(t('removedPerm')); } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); }
  loadPerms();
}
$('#shareForm').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const out = $('#shareResult'); out.hidden = true; out.classList.remove('err');
  const { valid, invalid } = parseEmails($('#shareEmails').value);
  const say = (txt, bad) => { out.textContent = txt; out.hidden = false; out.classList.toggle('err', !!bad); };
  if (!valid.length && !invalid.length) return say(t('noEmails'), true);
  if (invalid.length) return say(t('badEmails', { list: invalid.join(', ') }), true);
  const btn = $('#shareSend'); btn.disabled = true;
  try {
    const res = await state.project.share.emails(valid, $('#shareMsg').value.trim());
    const ok = res.filter(r => r.ok), bad = res.filter(r => !r.ok);
    const parts = [];
    if (ok.length) parts.push(t('sentOk', { n: ok.length }));
    if (bad.length) parts.push(t('sentFail', { list: bad.map(b => b.email).join(', ') }));
    say(parts.join(' '), bad.length > 0);
    if (!bad.length) $('#shareEmails').value = '';
    await loadPerms();
  } finally { btn.disabled = false; }
});
async function setLink(on) {
  try {
    const r = await state.project.share.setLink(on);
    toast(!on ? t('anyoneOff') : r && r.fallback ? t('orgFallback') : t('anyoneOn'));
  } catch (e) { toast(t('apiErr', { msg: errMsg(e) })); }
  await loadPerms();
}
$('#makeLink').addEventListener('click', () => setLink(true));
$('#stopLink').addEventListener('click', () => setLink(false));
$('#copyLink').addEventListener('click', async () => {
  const v = $('#linkInput').value;
  try { await navigator.clipboard.writeText(v); } catch (e) { $('#linkInput').select(); try { document.execCommand('copy'); } catch (e2) { /* ignore */ } }
  toast(t('copied'));
});

let qrLib;
async function drawQr(text) {
  const box = $('#qrBox');
  if (!text) { box.replaceChildren(); return; }
  try {
    if (!qrLib) {
      await new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js';
        s.integrity = 'sha384-8FWZA6BGMXhsfO+BLtrJK0We6gg5o1JyO8xQm6peWDEUs17ACA5ziE/NIAkl9z2k'; s.crossOrigin = 'anonymous'; // pinned file; CSP allows exactly this URL
        s.onload = res; s.onerror = rej; document.head.appendChild(s);
      });
      qrLib = window.qrcode;
    }
    const qr = qrLib(0, 'M'); qr.addData(text); qr.make();
    const n = qr.getModuleCount(), q = 2;
    let path = '';
    for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (qr.isDark(r, c)) path += `M${c + q} ${r + q}h1v1h-1z`;
    box.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n + q * 2} ${n + q * 2}" shape-rendering="crispEdges" role="img" aria-label="QR"><rect width="100%" height="100%" fill="#fff"/><path d="${path}" fill="#000"/></svg>`;
  } catch (e) { box.textContent = t('qrFail'); }
}

window.addEventListener('beforeunload', (e) => {
  if (state.tasks.some(x => ['uploading', 'retrying'].includes(x.state))) { e.preventDefault(); e.returnValue = ''; }
});

// MSAL popup return: the popup loads this page with the auth response in the hash; let MSAL consume it.
const inAuthPopup = isAuthPopup({ opener: window.opener, hash: location.hash, name: window.name });
if (!inAuthPopup) {
  applyLang(); show('signin');
  (async () => {
    try {
      await initCloud();
      renderChooser();
      const u = await cloud.restore(); // same-tab session from another Atelier tool, no popup
      if (u) await enterApp(u);
    } catch (e) { /* stay on the chooser */ }
  })();
}
