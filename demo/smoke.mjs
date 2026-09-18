// Smoke test for the local-mode demo (index.html served by demo/server.mjs).
//
// Runs the full no-server round-trip for ONE editor type:
//
//     node smoke.mjs [word|cell|slide]     (or: TYPE=cell node smoke.mjs)
//
// Verifies, per type:
//   - editor boots, document opens via loadBinary (seeded from the empty
//     template the demo server extracts out of the sdkjs sources),
//     onDocumentReady fires
//   - auth handshake answered by LocalMockServer (asc_isOffline, ServerIdWaitComplete)
//   - TWO marker edits (the block locks granted by LocalMockServer must let
//     the same block be modified again, see the lock checks below), Save ->
//     .bin persisted in localStorage, reload -> both edits survive, delete ->
//     storage reseeded from the template
//
// The lock checks (cell/slide only, word text edits take no block locks):
// the editors request element locks even in single-user mode (the edit is
// applied optimistically, the getLock reply updates the lock table). The
// reply MUST be attributed to the requester's composite
// `<idOriginal><indexUser>` id, otherwise the editor considers its own
// blocks locked by someone else and discards the subsequent edits of the
// same block (only one modifiable cell / one typed character per text box).
// Verified here by re-running the same lock gate the editor uses:
//   - cell: `api._isLockedAddWorksheets()` (the gate before adding a
//     worksheet),
//   - slide: `Document_Is_SelectionLocked(changestype_AddShape, <shape>)`
//     (the gate before adding/modifying a shape),
// both must keep answering "allowed" after a first edit already took the
// lock (with the composite-id bug the second call answers "locked").
import { chromium } from 'playwright';

const TYPE = (process.argv[2] || process.env.TYPE || 'word').toLowerCase();
const BASE = (process.env.BASE || 'http://localhost:4013') + '/?type=' + TYPE;
const LS_KEY = 'onlyoffice-local-doc-' + TYPE + '.bin';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function waitFor(fn, timeout, msg) {
    const start = Date.now();
    while (Date.now() - start < timeout) { const v = await fn(); if (v) return v; await sleep(100); }
    throw new Error('timeout waiting for ' + msg);
}
function check(cond, msg) { if (!cond) { console.error('FAIL:', msg); process.exit(1); } console.log('ok -', msg); }

// --- page-side helpers (run inside the demo page via page.evaluate) ---
// Resolving the editor api: word/slide expose it as `window.editor`, the
// cell build as `window.editorCell` (same logic as the wrapper's getApi).

function pageInsert(type, marker) {
    return `(() => {
        const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
        const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
            .find((c) => c && typeof c.asc_Save === 'function');
        if (!api) return false;
        if (${JSON.stringify(type)} === 'word') api.asc_AddText(${JSON.stringify(marker)});
        else if (${JSON.stringify(type)} === 'cell') api.asc_addWorksheet(${JSON.stringify(marker)});
        else api.AddSlide();
        return true;
    })()`;
}

// Returns the marker-independent "size" of the document (word: text length,
// cell: worksheet count, slide: slide count) — used as a baseline after
// reseeding, and to observe the edits.
function pageCount(type) {
    return `(() => {
        const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
        const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
            .find((c) => c && typeof c.asc_Save === 'function');
        if (!api) return null;
        if (${JSON.stringify(type)} === 'word') {
            const doc = api.WordControl && api.WordControl.m_oLogicDocument;
            return doc && typeof doc.GetText === 'function' ? doc.GetText().length : 0;
        }
        if (${JSON.stringify(type)} === 'cell') {
            return typeof api.asc_getWorksheetsCount === 'function' ? api.asc_getWorksheetsCount() : 0;
        }
        return api.WordControl && typeof api.WordControl.GetSlidesCount === 'function'
            ? api.WordControl.GetSlidesCount() : 0;
    })()`;
}

// True once ALL the marker edits are present in the model.
function pagePresentAll(type, markers, slideBaseline) {
    return `(() => {
        const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
        const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
            .find((c) => c && typeof c.asc_Save === 'function');
        if (!api) return false;
        const markers = ${JSON.stringify(markers)};
        if (${JSON.stringify(type)} === 'word') {
            const doc = api.WordControl && api.WordControl.m_oLogicDocument;
            const t = doc && typeof doc.GetText === 'function' ? doc.GetText() : '';
            return markers.every((m) => t.indexOf(m) !== -1);
        }
        if (${JSON.stringify(type)} === 'cell') {
            const n = typeof api.asc_getWorksheetsCount === 'function' ? api.asc_getWorksheetsCount() : 0;
            const names = [];
            for (let i = 0; i < n; i++) {
                const ws = api.wbModel && api.wbModel.getWorksheet(i);
                if (ws && ws.getName) names.push(ws.getName());
            }
            return markers.every((m) => names.indexOf(m) !== -1);
        }
        // slide: the edits were added slides (one per marker).
        const baseline = ${JSON.stringify(slideBaseline || 0)};
        return api.WordControl && typeof api.WordControl.GetSlidesCount === 'function'
            && api.WordControl.GetSlidesCount() >= baseline + markers.length;
    })()`;
}

// The state of the co-authoring client's lock table: `settled` once every
// requested lock has been answered by the mock server (state 1 = "asked",
// 2 = "acquired by me", 3 = "held by someone else").
const PAGE_LOCK_STATE = `(() => {
    const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
    const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
        .find((c) => c && typeof c.asc_Save === 'function');
    if (!api) return null;
    const coApi = api.CoAuthoringApi && api.CoAuthoringApi._CoAuthoringApi;
    if (!coApi || !coApi._locks) return null;
    const states = Object.keys(coApi._locks).map((k) => coApi._locks[k].state);
    return { settled: states.length > 0 && states.every((s) => s !== 1), states: states };
})()`;

// Re-runs the same lock gate the editor uses before an edit of the same
// block, and returns true when the edit would be ALLOWED (false = the block
// is considered locked, by someone else). Word has no gate to re-run (its
// text edits take no block locks): returns null.
function pageLockGate(type) {
    if (type === 'word') {
        return null;
    }
    if (type === 'cell') {
        // The gate before adding a worksheet (same block every time): this is
        // exactly what `_isLockedAddWorksheets` runs internally — NB it cannot
        // be used directly here since it drops `lock()`'s return value.
        return `(() => {
            const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
            const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
                .find((c) => c && typeof c.asc_Save === 'function');
            const ACE = w.AscCommonExcel;
            if (!api || !api.collaborativeEditing || !ACE || !ACE.c_oAscLockTypeElem) return null;
            const lockInfo = api.collaborativeEditing.getLockInfo(
                ACE.c_oAscLockTypeElem.Sheet, null, ACE.c_oAscLockAddSheet, ACE.c_oAscLockAddSheet);
            // true = the edit is allowed (mine or pending), false = the block
            // is considered locked by someone else.
            return api.collaborativeEditing.lock([lockInfo], function () {});
        })()`;
    }
    // slide: the gate before adding/modifying a shape, re-run for a shape of
    // the (last) added slide.
    return `(() => {
        const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
        const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
            .find((c) => c && typeof c.asc_Save === 'function');
        if (!api || !w.AscCommon || !w.AscCommon.changestype_AddShape) return null;
        const doc = api.WordControl && api.WordControl.m_oLogicDocument;
        if (!doc || typeof doc.Document_Is_SelectionLocked !== 'function') return null;
        const slide = doc.Slides[doc.Slides.length - 1];
        const shapes = (slide && slide.cSld && slide.cSld.spTree) || [];
        const shape = shapes.find((s) => s && s.Lock && s.Get_Id);
        if (!shape) return null;
        // true = locked (the edit would be discarded).
        return !doc.Document_Is_SelectionLocked(w.AscCommon.changestype_AddShape, shape);
    })()`;
}

const browser = await chromium.launch({ args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage'] });
const page = await browser.newPage();
page.on('pageerror', (e) => console.log('pageerror:', String(e).slice(0,300)));
page.on('console', (m) => { if (m.type() === 'error') console.log('console.error:', m.text().slice(0,200)); });

await page.goto(BASE, { waitUntil: 'domcontentloaded', timeout: 30000 });
await page.locator('iframe[name="frameEditor"]').waitFor({ timeout: 30000 });
console.log('ok - editor iframe created (' + TYPE + ')');

// Document ready (proves loadBinary + LocalMockServer auth handshake both worked).
await waitFor(() => page.evaluate(() => window.document.getElementById('status').textContent === 'document ready'), 120000, 'document ready');
console.log('ok - document ready (loadBinary + LocalMockServer auth)');

// Local flags + baseline size of the freshly seeded document.
const flags = await page.evaluate((expr) => eval(expr), pageCount(TYPE));
const offline = await page.evaluate(() => {
    const w = document.querySelector('iframe[name="frameEditor"]').contentWindow;
    const api = [w.editor, w.editorCell, w.Asc && w.Asc.editor]
        .find((c) => c && typeof c.asc_Save === 'function');
    return {
        offlineFlag: !!w.__OOCryptPadOffline,
        isOffline: api && typeof api.asc_isOffline === 'function' ? api.asc_isOffline() : undefined,
    };
});
console.log('flags:', JSON.stringify(offline), 'baseline count:', flags);
check(offline.offlineFlag === true, '__OOCryptPadOffline flag set (local mode)');
check(offline.isOffline === true, 'asc_isOffline() === true');
check(typeof flags === 'number', 'document model reachable (baseline ' + flags + ')');

// Marker edits. NB: for cell the markers double as the worksheet names, and
// sheet names are capped at 31 chars (longer names silently fall back to
// "SheetN"), so keep them short there.
const MARKERS = TYPE === 'cell'
    ? ['LC1_' + Date.now(), 'LC2_' + Date.now()]
    : ['LOCAL_DEMO_MARKER1_' + TYPE + '_' + Date.now(), 'LOCAL_DEMO_MARKER2_' + TYPE + '_' + Date.now()];

// --- edit #1 ---
const inserted1 = await page.evaluate((expr) => eval(expr), pageInsert(TYPE, MARKERS[0]));
check(inserted1 === true, 'insert edit #1 via editor api (' + TYPE + ')');
await waitFor(() => page.evaluate((expr) => eval(expr), pagePresentAll(TYPE, [MARKERS[0]], flags)), 30000, 'marker #1 present after edit');
console.log('ok - marker #1 present in document after edit');

// --- lock regression checks: the block lock taken for an edit must be
// recognized as OURS, so the same kind of edit is allowed again. With the
// composite-id bug, the getLock reply marks our own blocks as foreign-locked
// and the subsequent edits of the same block are discarded (only one
// modifiable cell / one typed character per text box).
if (TYPE === 'cell') {
    // The worksheet-add (edit #1) already requested the add-sheet block lock:
    // the reply must have been attributed to us (state 2 = acquired by me;
    // state 3 = held by someone else).
    const lockInfo = await waitFor(async () => {
        const st = await page.evaluate((expr) => eval(expr), PAGE_LOCK_STATE);
        return st && st.settled ? st : null;
    }, 30000, 'lock reply');
    check(lockInfo.settled === true, 'getLock request answered by LocalMockServer');
    check(lockInfo.states.every((s) => s === 2), 'lock granted to us (state 2 ' + JSON.stringify(lockInfo.states) + ', not 3/another user)');
    // Re-run the same gate the editor uses before adding a worksheet: with
    // the bug the poisoned block makes it answer "locked".
    const gateExpr = pageLockGate(TYPE);
    const gate1 = await page.evaluate((expr) => eval(expr), gateExpr);
    check(gate1 === true, 'lock gate allows the edit (cell)');
    const gate2 = await page.evaluate((expr) => eval(expr), gateExpr);
    check(gate2 === true, 'lock gate still allows the same-block edit (cell, no self-poisoning)');
} else if (TYPE === 'slide') {
    // AddSlide takes no block lock: run the gate the editor uses before
    // adding/modifying a shape, which both checks and requests the lock.
    const gateExpr = pageLockGate(TYPE);
    const gate1 = await page.evaluate((expr) => eval(expr), gateExpr);
    check(gate1 === true, 'lock gate allows the edit (slide)');
    // The gate just requested the shape lock: the reply must have been
    // attributed to us (state 2 = acquired by me; state 3 = held by someone
    // else).
    const lockInfo = await waitFor(async () => {
        const st = await page.evaluate((expr) => eval(expr), PAGE_LOCK_STATE);
        return st && st.settled ? st : null;
    }, 30000, 'lock reply');
    check(lockInfo.settled === true, 'getLock request answered by LocalMockServer');
    check(lockInfo.states.every((s) => s === 2), 'lock granted to us (state 2 ' + JSON.stringify(lockInfo.states) + ', not 3/another user)');
    // Re-run the same gate: with the bug the first reply marked the shape as
    // foreign-locked, so this second call answers "locked".
    const gate2 = await page.evaluate((expr) => eval(expr), gateExpr);
    check(gate2 === true, 'lock gate still allows the same-block edit (slide, no self-poisoning)');
}

// --- edit #2 (the same kind of edit again: for the lock-owner bug this one
// would be discarded) ---
const inserted2 = await page.evaluate((expr) => eval(expr), pageInsert(TYPE, MARKERS[1]));
check(inserted2 === true, 'insert edit #2 via editor api (' + TYPE + ')');
await waitFor(() => page.evaluate((expr) => eval(expr), pagePresentAll(TYPE, MARKERS, flags)), 30000, 'both markers present after edits');
console.log('ok - both markers present in document after the two edits');

// Save -> localStorage.
await page.click('#btn-save');
await waitFor(() => page.evaluate(() => {
    const s = document.getElementById('status').textContent;
    return s.startsWith('saved') || s.startsWith('save error');
}), 30000, 'save complete');
const saveStatus = await page.evaluate(() => document.getElementById('status').textContent);
check(/saved .* bytes/.test(saveStatus), 'save produced bytes: ' + saveStatus);
const storedLen = await page.evaluate((key) => localStorage.getItem(key)?.length || 0, LS_KEY);
check(storedLen > 0, '.bin stored in localStorage (base64 len=' + storedLen + ')');

// Reload from storage -> both edits survive the save+reopen round-trip.
await page.click('#btn-reload');
await page.locator('iframe[name="frameEditor"]').waitFor({ timeout: 30000 });
await waitFor(() => page.evaluate(() => document.getElementById('status').textContent === 'document ready'), 120000, 'reloaded document ready');
const persisted = await waitFor(() => page.evaluate((expr) => eval(expr), pagePresentAll(TYPE, MARKERS, flags)), 60000, 'edits persisted after reload');
check(persisted === true, 'both edits persisted across save + reload');
console.log('ok - both edits persisted across save/reload round-trip');

// Delete -> localStorage cleared, editor reseeded from the empty template.
await page.click('#btn-delete');
await page.locator('iframe[name="frameEditor"]').waitFor({ timeout: 30000 });
await waitFor(() => page.evaluate(() => document.getElementById('status').textContent === 'document ready'), 120000, 'reseeded document ready');
const afterCount = await page.evaluate((expr) => eval(expr), pageCount(TYPE));
check(afterCount === flags, 'document back to the seeded baseline after delete (' + afterCount + ')');
const storedAfter = await page.evaluate((key) => localStorage.getItem(key), LS_KEY);
check(!!storedAfter, 'localStorage reseeded after delete (empty template)');

await browser.close();
console.log('ALL PASS (' + TYPE + ')');
