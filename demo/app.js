// OnlyOffice "local" mode demo.
//
// Everything happens client-side: there is no DocumentServer, no co-authoring
// network, no `/downloadas` converter, no x2t. The document is stored as an
// OnlyOffice `Editor.bin` (`.bin`) in `localStorage` and (re)opened with
// `DocEditor.loadBinary()`. Pick the editor with `?type=word|cell|slide`
// (each build has its own `Editor.bin` signature: DOCY/XLSY/PPTY).
//
// The wrapper (`onlyoffice-editor/src/index.ts`) provides the API:
//
//   - `config.offline`            — local (no-server) mode.
//   - `config.autosave`         — 0 disables the SDK periodic autosave
//     entirely; non-zero only ENABLES it — the SDK tests the value solely
//     against zero (`apiBase._autoSave`), it is not used as an interval. The
//     effective cadence is hardcoded (`autoSaveGapFast`, 2 s for this
//     single-user editor; see applyAutosave in the wrapper).
//   - `config.events.save(bytes)` — asynchronous host save, called by the
//     wrapper inside the wrapped `api.asc_Save` (the host Save button, the
//     editor's own toolbar/File-menu Save and autosave all go through it)
//     with the raw `Editor.bin` bytes.
//   - `config.events.printPdf(dc, cb)` / `config.events.downloadAs(dc, cb)`
//     — bridged by the wrapper to the `window.APP` host hooks the editor calls
//     for Print / Download-as instead of POSTing to `/downloadas`.
//
// Editor lifecycle (see start()): `new DocEditor(placeholder)` (creates
// the `window.APP` host-hook surface — the wrapper plays the host here) ->
// `await init(config, mockServer)` (bridges the Print/Download-as hooks onto
// it, inserts the iframe, connects the mock server, resolves on onAppReady)
// -> `sendMessageToOO` (license) -> `loadBinary(bytes)`.
//
// Save model (mirrors Cryptpad's embedded/integration mode): the OO config
// registers NO `onSaveDocument` handler, so `checkSaveDocumentEvent` is inert
// (it would otherwise POST to `/downloadas`); the wrapper serialises via the
// editor's own `asc_nativeGetFile()` (the call CryptPad's `getContent()`
// uses) and hands the bytes to `config.events.save`.
// The co-authoring "server" side of the local mode is provided below by the
// demo's own `LocalMockServer` (single-user `auth`/`saveChanges` answers).
//
// Print / Download-as: producing a real PDF/docx needs x2t (or the sdkjs-ooxml
// addon), out of scope here — the hooks end the editor action cleanly and the
// Download-as one downloads the `.bin` (the local format).

const $ = (id) => document.getElementById(id);
const status = (t) => { $("status").textContent = t; };

// ---- editor type (word | cell | slide) ----

// Every editor build is served from the same sdkjs tree: pick one with
// ?type=. Each build serialises to its own `Editor.bin` signature
// ("DOCY"/"XLSY"/"PPTY" — see the wrapper's decodeNativeFileString), so the
// stored document is keyed per type.
const TYPE = (new URLSearchParams(location.search).get("type") || "word").toLowerCase();
const TYPES = {
    word:  { documentType: "word",  fileType: "docx", title: "local.docx" },
    cell:  { documentType: "cell",  fileType: "xlsx", title: "local.xlsx" },
    slide: { documentType: "slide", fileType: "pptx", title: "local.pptx" },
};
if (!TYPES[TYPE]) status("unknown ?type=" + TYPE + " — using word");
const T = TYPES[TYPE] || TYPES.word;
const LS_KEY = "onlyoffice-local-doc-" + T.documentType + ".bin";

// ---- localStorage helpers (localStorage is string-only -> base64) ----

function bytesToBase64(bytes) {
    let bin = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(bin);
}

function base64ToBytes(b64) {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
}

function loadStored() {
    const b64 = localStorage.getItem(LS_KEY);
    return b64 ? base64ToBytes(b64) : null;
}

function storeBytes(bytes) {
    localStorage.setItem(LS_KEY, bytesToBase64(bytes));
}

function deleteStored() {
    localStorage.removeItem(LS_KEY);
}

// Fetch the empty-document `Editor.bin` seed for the current editor type
// from the demo server, which extracts it from the sdkjs sources (each
// editor build has its own signature: DOCY/XLSY/PPTY). Served as the
// "<signature>;v<version>;<length>;<base64>" wire-string bytes — the same
// form saves are stored in (the editor's open path requires the signature
// header).
async function fetchEmptyBin() {
    const res = await fetch("/empty.bin?type=" + T.documentType);
    if (!res.ok) throw new Error("could not fetch /empty.bin: " + res.status);
    return new Uint8Array(await res.arrayBuffer());
}

// ---- editor helpers ----

/** Trigger a browser download of `bytes` as `filename`. */
function downloadBytes(bytes, filename) {
    const blob = new Blob([bytes], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// ---- LocalMockServer ----

/**
 * A `MockServer` for the **local** mode (no server at all, everything happens
 * in the browser).
 *
 * The editor still runs its co-authoring state machine over the Gateway
 * message bridge, so it expects a "server" to answer the `auth` handshake and
 * the save-lock requests. `LocalMockServer` answers them locally so that:
 *
 *   - `onDocumentReady` fires: the `auth` response sets `ServerIdWaitComplete`
 *     (the gate in `_openDocumentEndCallback`). The document bytes themselves
 *     are opened by the host via `DocEditor.loadBinary()`;
 *   - the editor's `asc_Save()` save-lock request resolves promptly instead of
 *     stalling ~10 s on the `askSaveChanges` timeout.
 *
 * It is a single-user, no-changes server: there are no other participants and
 * the initial change set is empty. The document is supplied entirely by the
 * host's `loadBinary()` call, so there is no `documentOpen` / `Editor.bin` URL
 * here.
 */
class LocalMockServer {
    constructor(editor, options) {
        this.editor = editor;
        const u = (options && options.user) || { id: "uid-local", username: "Local user" };
        this.userId = u.id;
        this.username = u.username;
        // Server-assigned participant index (see answerAuth). The OnlyOffice
        // wire protocol identifies users by the *composite*
        // `<idOriginal><indexUser>` id: on `auth` DocsCoApi sets
        // `_userId = user.id + indexUser` and every lock/change record's
        // `user` field uses that composite (see `DocsCoApi._onGetLock` and the
        // captured real session in the parsec-cloud docs,
        // docs/rfcs/1030-collaborative-editics/oo_example_session.md).
        this.indexUser = 1;
        this.compositeUserId = this.userId + this.indexUser;
        // Monotonic change counter (the OnlyOffice "puckerIndex"): the total
        // number of changes the server has stored. Used to answer the
        // saveChanges flow (`savePartChanges`/`unSaveLock`).
        this.syncChangesIndex = 0;
        // Absolute change index at which the in-progress save starts (the
        // "save point" reported back in the final `unSaveLock`).
        this.saveChangesIndex = 0;
    }

    onMessage(msg) {
        switch (msg && msg.type) {
            case "auth":
                this.answerAuth(msg);
                break;
            case "isSaveLock":
                // Grant the save lock immediately so `asc_Save`'s
                // `askSaveChanges` resolves without the ~10 s timeout.
                this.editor.sendMessageToOO({ type: "saveLock", saveLock: false });
                break;
            case "unSaveLock":
                // Cancellation of an in-progress save: release the save lock
                // and acknowledge with -1 indices.
                this.editor.sendMessageToOO({ type: "unSaveLock", index: -1, time: -1, syncChangesIndex: -1 });
                break;
            case "saveChanges":
                // Local mode persists via the binary-save path
                // (`asc_onSaveDocument`), not via the co-authoring channel: only
                // drive the client's save state machine (chunk acks + final
                // unSaveLock).
                this.answerSaveChanges(msg);
                break;
            case "getLock":
                // Element locks (cell/slide editors ask for these, e.g. while
                // adding a worksheet, whenever the collaborative state is
                // not "single user" — such as right after document ready).
                // Without a reply the request times out and the edit is
                // silently dropped. They MUST be granted to the composite user
                // id (see answerGetLock).
                this.answerGetLock(msg);
                break;
            case "unLockDocument":
                // Fire-and-forget cleanup (there is no other participant to
                // notify and no auth lock to manage in single-user local
                // mode); only an in-progress save needs an answer.
                if (msg.isSave) {
                    this.editor.sendMessageToOO({ type: "unSaveLock", index: -1, time: -1, syncChangesIndex: -1 });
                }
                break;
            default:
                // Other message types (cursor, meta, authChangesAck, ...) are
                // not relevant in local single-user mode: ignore them.
                break;
        }
    }

    /**
     * Grant every requested block to the requester (single-user server).
     *
     * Mirrors the DocumentServer's `getLock` reply, the shape the editor's
     * `DocsCoApi._onGetLock` expects: `locks` maps each block id to
     * `{user, time, block}`, where `user` MUST be the requester's composite
     * `<idOriginal><indexUser>` id for the lock to be recognized as the
     * requester's own (state 2, "acquired by me"); `block` is the original
     * block descriptor — the cell/slide editors read `lock["block"]["guid"]`
     * as the block id (the word editor sends plain guid strings).
     *
     * NB: granting with the plain user id instead of the composite one makes
     * the editor consider its own blocks as locked by someone else (state 3
     * -> `onLocksAcquired` marks them foreign-locked), which silently discards
     * subsequent edits of the same block: the second character typed in a
     * text box, the second modification of the same cell, ...
     */
    answerGetLock(msg) {
        const blocks = (msg && msg.block) || [];
        const locks = {};
        for (const b of blocks) {
            const key = typeof b === "object" && b !== null ? b.guid : b;
            if (key !== undefined) {
                locks[key] = {
                    user: this.compositeUserId,
                    time: typeof b === "object" && b !== null && b.time !== undefined ? b.time : Date.now(),
                    block: b,
                };
            }
        }
        this.editor.sendMessageToOO({ type: "getLock", locks });
    }

    /**
     * Drive the client's save state machine like the DocumentServer does:
     * intermediate chunks get a `savePartChanges` ack (the client emits the
     * next chunk on it), the final chunk gets an `unSaveLock` carrying the
     * "save point" change index, the save time and the new total change
     * count.
     */
    answerSaveChanges(msg) {
        if (msg.startSaveChanges) {
            // The save point of this save is the change index it starts at
            // (the previous total).
            this.saveChangesIndex = this.syncChangesIndex;
        }
        this.syncChangesIndex += this.countChanges(msg);
        if (msg.endSaveChanges) {
            this.editor.sendMessageToOO({
                type: "unSaveLock",
                index: this.saveChangesIndex,
                time: Date.now(),
                syncChangesIndex: this.syncChangesIndex,
            });
        } else {
            // Acknowledge the chunk so the client emits the next one (the -1
            // index leaves the client's save point unchanged).
            this.editor.sendMessageToOO({
                type: "savePartChanges",
                changesIndex: msg.startSaveChanges ? this.saveChangesIndex : -1,
                syncChangesIndex: this.syncChangesIndex,
            });
        }
    }

    // Counts the changes in a saveChanges chunk: a JSON-serialized array by
    // default (the binary form is only used with `binaryChanges`).
    countChanges(msg) {
        try {
            let changes = msg.changes;
            if (typeof changes === "string") {
                changes = JSON.parse(changes);
            }
            return Array.isArray(changes) ? changes.length : 0;
        } catch (_e) {
            return 0;
        }
    }

    /**
     * Answer the `auth` handshake.
     *
     * Mirrors the shape the real DocumentServer sends and that the editor's
     * `DocsCoApi._onAuth` expects: `authChanges` first, then `auth` with
     * `result:1`, the participants list and the user's `indexUser`. `result:1`
     * is mandatory — without it `_onAuth` never sets `_isAuth`, so
     * `onFirstLoadChangesEnd` -> `asyncServerIdEndLoaded` never runs and
     * `ServerIdWaitComplete` stays false, which blocks `onDocumentReady`.
     * A participant is identified by its composite `<idOriginal><indexUser>`
     * id.
     *
     * No `documentOpen` is sent: the document is opened by the host via
     * `DocEditor.loadBinary()`. No `connectState` either: without it the
     * editor stays in its single-user state, where edits are applied
     * optimistically and element locks are asked fire-and-forget (a
     * connectState with a single participant would flip it to the "exclusive
     * editor" state, which starts saving on its own).
     */
    answerAuth(_authMsg) {
        const participant = {
            id: this.compositeUserId,
            idOriginal: this.userId,
            username: this.username,
            indexUser: this.indexUser,
            connectionId: "conn-local",
            isCloseCoAuthoring: false,
            view: false,
        };
        const changes = [];

        this.editor.sendMessageToOO({ type: "authChanges", changes });

        this.editor.sendMessageToOO({
            type: "auth",
            result: 1,
            sessionId: "local-session",
            participants: [participant],
            locks: [],
            changes,
            changesIndex: 0,
            indexUser: this.indexUser,
            buildVersion: "5.2.6",
            buildNumber: 2,
            licenseType: 3,
        });
    }
}

// ---- the demo ----

let docEditor = null;
let mockServer = null;

function setButtons(enabled) {
    $("btn-save").disabled = !enabled;
    $("btn-reload").disabled = !enabled;
    $("btn-delete").disabled = !enabled;
}

function makeEditorConfig() {
    return {
        offline: true,
        // Autosave: 0 disables the SDK's periodic autosave entirely (manual
        // save only in this demo). NB: the SDK does NOT use this value as an
        // interval — `apiBase._autoSave` only tests `0 !== autoSaveGap`. With
        // a non-zero value the effective cadence for this single-user editor
        // is the hardcoded `autoSaveGapFast` (2 s after the last change), not
        // the configured number. Applied by the wrapper via
        // `api.asc_setAutoSaveGap` on document ready (see applyAutosave).
        autosave: 0,
        width: "100%",
        height: "100%",
        documentType: T.documentType,
        document: {
            key: "local-doc-" + T.documentType + "-key",
            title: T.title,
            fileType: T.fileType,
            // A truthy placeholder URL. The vanilla api-orig.js _checkConfigParams
            // rejects an empty/falsy `url` (alerts and aborts editor creation),
            // so it cannot be "". It must NOT be one of the SDK's magic sentinels
            // ("_offline_"/"_chart_"/"_ole_"/"_data_"), which would trigger the
            // desktop-only empty-doc crash. And since LocalMockServer never sends
            // `documentOpen`, this URL is never fetched — it only needs to pass
            // validation. The document bytes come from `loadBinary()`.
            url: "local://document",
            // Print and Download-as are enabled: the wrapper bridges them to
            // `config.events.printPdf` / `config.events.downloadAs` (see
            // sdkjs/common/apiBase.js) instead of POSTing to `/downloadas`.
            permissions: { edit: true, download: true, print: true },
        },
        editorConfig: {
            mode: "edit",
            customization: {},
            user: { id: "uid-local", name: "Local user" },
        },
        events: {
            onAppReady: () => { status("app ready"); },
            onDocumentReady: () => {
                status("document ready");
                setButtons(true);
            },
            // NOTE: no `onSaveDocument` handler. Registering one would set
            // `canSaveDocumentToBinary` -> `DocInfo.SupportsOnSaveDocument` true
            // -> `checkSaveDocumentEvent` would run and POST the `.bin` to
            // `/downloadas` (404 -> hung spinner). Leaving it unset makes that
            // path inert; save is driven by the wrapper's `asc_Save` hook below.
            //
            // Asynchronous host save, called by the wrapper inside the wrapped
            // `api.asc_Save` with the raw `Editor.bin` bytes (serialized in-browser
            // via the editor's own `asc_nativeGetFile()` — no server). The host
            // Save button, the editor's own toolbar/File-menu Save, and autosave
            // all go through it.
            onSave: async (bytes) => {
                storeBytes(bytes);
                status("saved " + bytes.byteLength + " bytes to local storage");
            },
            // Print button. `dataContainer.data` is the OnlyOffice renderer
            // binary (NOT a PDF); a real PDF needs x2t — out of scope. `cb(null)`
            // ends the editor action cleanly (no "Unknown error", no spinner).
            onPrintPdf: (dataContainer, cb) => {
                status("Print needs x2t PDF conversion (out of scope for local mode)");
                cb(null);
            },
            // Download-as / Save-copy. `dataContainer.data` is an `Editor.bin`
            // in this build (NOT a real docx — needs x2t/sdkjs-ooxml, out of
            // scope), so we download the `.bin` (the local format).
            onDownloadAs: (dataContainer, cb) => {
                const bytes = dataContainer && dataContainer.data;
                if (bytes && bytes.byteLength) {
                    downloadBytes(bytes, "document.bin");
                    status("downloaded document.bin (.bin is the local format; x2t out of scope)");
                } else {
                    status("download-as produced no bytes");
                }
                cb(null);
            },
            onError: (e) => { status("error: " + JSON.stringify(e && e.data)); },
        },
    };
}

async function start() {
    setButtons(false);

    // 1. Create the editor and initialise it with the config.

    docEditor = new DocsAPI.DocEditor("placeholder");
    mockServer = new LocalMockServer(docEditor);
    await docEditor.init(makeEditorConfig(), mockServer);

    // 2. Push the license into the editor: without it the editor idles in
    //    WaitAuth with no permissions and never opens the document.

    docEditor.sendMessageToOO(
        {
            type: 'license',
            license: {
                type: 3,
                mode: 0,
                // light: false,
                // trial: false,
                rights: 1,
                buildVersion: "7.3.3",
                buildNumber: 8,
                // branding: false
            }
        }
    );

    // 3. Open the document by pushing its bytes: from local storage, or —
    //    first run — the empty-document template the demo server extracts
    //    from the sdkjs sources for this ?type= editor (DOCY/XLSY/PPTY).
    let bytes = loadStored();
    if (!bytes) {
        status("seeding empty " + T.documentType + " document…");
        bytes = await fetchEmptyBin();
        storeBytes(bytes);
        status("seeded empty " + T.documentType + " document into local storage");
    } else {
        status("loaded .bin from local storage (" + bytes.byteLength + " bytes)");
    }

    status("opening document via loadBinary…");
    // Keep a copy because loadBinary transfers (neuters) the buffer.
    docEditor.loadBinary(bytes.slice());
}

function destroyEditor() {
    if (docEditor) {
        try { docEditor.destroyEditor(); } catch (e) {}
        docEditor = null;
        mockServer = null;
    }
    const old = document.querySelector('iframe[name="frameEditor"]');
    if (old) old.remove();
}

async function save() {
    setButtons(false);
    status("serializing to .bin…");
    // Goes through the wrapper's wrapped `api.asc_Save`:
    //   serializeBinary -> config.events.save(bytes) -> SDK save handshake.
    // The `config.events.save` handler above stores and updates the status.
    docEditor.save();
    setButtons(true);
}

async function reload() {
    setButtons(false);
    const bytes = loadStored();
    if (!bytes) { status("nothing stored to reload"); setButtons(true); return; }
    status("reloading from local storage…");
    destroyEditor();
    await start();
}

async function deleteDoc() {
    setButtons(false);
    deleteStored();
    status("deleted .bin from local storage; reloading fresh…");
    destroyEditor();
    await start();
}

// Wire up the buttons.
$("btn-save").addEventListener("click", save);
$("btn-reload").addEventListener("click", reload);
$("btn-delete").addEventListener("click", deleteDoc);

// Boot. `DocsAPI` is provided by the wrapper script loaded in index.html.
start().catch((e) => status("boot error: " + e));
