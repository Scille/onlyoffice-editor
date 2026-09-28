import { EventHandler, HandlerHandle } from "./eventHandler";
import { deepAssign, noop, waitForEvent } from "./utils";

let DocEditorOrig: any;

// ---------------------------------------------------------------------------
// OnlyOffice protocol
// ---------------------------------------------------------------------------

/** A co-authoring participant, as carried by `auth`/`waitAuth`/`connectState`. */
export interface OOParticipantEntry {
    /** Composite `<idOriginal><indexUser>` id (the editor's own `_userId`). */
    id: string;
    /** Integrator-provided user id. */
    idOriginal: string;
    /** Display name. */
    username: string;
    /** Participant index (order of arrival). */
    indexUser: number;
    /** Read-only viewer. */
    view: boolean;
    /** Underlying connection id (= sessionId). */
    connectionId?: string;
    isCloseCoAuthoring?: boolean;
    isLiveViewer?: boolean;
    encrypted?: boolean;
}

/** A chat/comment message entry (server `message`/`auth` events). */
export interface OOMessageEntry {
    docid: string;
    message: string;
    time: number;
    user: string;
    useridoriginal: string;
    username: string;
}

/**
 * An element-lock block descriptor: a plain guid string (word) or an object
 * with a `guid` (cell/slide/pdf).
 */
export interface OOBlockDescriptor {
    guid?: string;
    time?: number;
    [key: string]: unknown;
}

/** An element lock granted to a user (server `getLock`/`auth` events). */
export interface OOLockEntry {
    /** Composite user id of the lock holder. */
    user: string;
    time: number;
    /** The block descriptor the lock was requested for. */
    block: unknown;
}

/** A serialized change entry (server `authChanges`/`saveChanges` events). */
export interface OOChangeEntry {
    docid: string;
    /** JSON-encoded opaque op fragment. */
    change: string;
    time: number;
    user: string;
    useridoriginal: string;
}

/** A lock released with a `saveChanges`/`releaseLock` event. */
export interface OOSavedLockEntry {
    block: unknown;
    user: string;
    time: number;
    changes: unknown;
}

/** Connection limits/behaviour knobs carried by the server `auth` event. */
export interface OOServerAuthSettings {
    spellcheckerUrl?: string;
    reconnection?: { attempts: number; delay: number };
    /** Binary (non-JSON) `saveChanges` encoding. */
    binaryChanges?: boolean;
    websocketMaxPayloadSize?: number;
    maxChangesSize?: number;
    limits_image_size?: number;
    limits_image_types_upload?: string;
}

// --- Client events (editor -> server) ------------------------------------

/** Editor auth handshake; the "server" must answer with a `ToOO` `auth`. */
export interface OOClientEventAuth {
    type: "auth";
    docid: string;
    token: string;
    user: {
        id: string;
        username: string;
        firstname: string | null;
        lastname: string | null;
        indexUser: number;
    };
    editorType: number;
    lastOtherSaveTime: number;
    block: unknown[];
    sessionId: string | null;
    sessionTimeConnect: number | null;
    sessionTimeIdle: number;
    documentFormatSave: number;
    isCloseCoAuthoring: boolean;
    openCmd: Record<string, unknown> | null;
    lang: string;
    mode: string;
    permissions: { edit: boolean; review: boolean };
    encrypted: boolean;
    IsAnonymousUser: boolean;
    timezoneOffset: number;
    headingsColor: string | null;
    coEditingMode: string;
    jwtOpen: string;
    jwtSession?: string;
    time: number;
    supportAuthChangesAck: boolean;
}

/** A chat/comment message broadcast to the co-editors. */
export interface OOClientEventMessage {
    type: "message";
    message: string;
}

/** A cursor position broadcast (opaque OnlyOffice-internal string). */
export interface OOClientEventCursor {
    type: "cursor";
    cursor: string;
}

/** Element-lock request (cell ranges, slide objects, ...). */
export interface OOClientEventGetLock {
    type: "getLock";
    block: (OOBlockDescriptor | string)[];
}

/** Save-lock request. */
export interface OOClientEventIsSaveLock {
    type: "isSaveLock";
    syncChangesIndex: number;
}

/** A chunk of serialized document changes to persist. */
export interface OOClientEventSaveChanges {
    type: "saveChanges";
    /**
     * JSON-encoded string of op fragments (default mode) or a real array
     * (binary-changes mode).
     */
    changes: string | unknown[];
    startSaveChanges: boolean;
    endSaveChanges: boolean;
    isCoAuthoring?: boolean;
    isExcel?: boolean;
    deleteIndex?: number | null;
    excelAdditionalInfo?: string | null;
    unlock?: boolean;
    releaseLocks?: boolean;
    reSave?: number;
}

/** Save-lock release (cancellation of an in-progress save). */
export interface OOClientEventUnSaveLock {
    type: "unSaveLock";
}

/** Document unlock (end of an exclusive editing session). */
export interface OOClientEventUnLockDocument {
    type: "unLockDocument";
    isSave: boolean;
    unlock: boolean;
    deleteIndex?: number | null;
    releaseLocks?: boolean;
}

/** Connection close. */
export interface OOClientEventClose {
    type: "close";
}

/** Ack of the server `authChanges` event. */
export interface OOClientEventAuthChangesAck {
    type: "authChangesAck";
}

/** Request the missed chat/comment messages. */
export interface OOClientEventGetMessages {
    type: "getMessages";
}

/** Ask the "server" to open the document (unused offline: `loadBinary`). */
export interface OOClientEventOpenDocument {
    type: "openDocument";
    message: Record<string, unknown>;
}

/** Client-side log entry. */
export interface OOClientEventClientLog {
    type: "clientLog";
    level: string;
    msg: string;
}

/** Session keep-alive. */
export interface OOClientEventExtendSession {
    type: "extendSession";
    idletime: number;
}

/** A force save has started. */
export interface OOClientEventForceSaveStart {
    type: "forceSaveStart";
}

/** RPC roundtrip initiated by the editor. */
export interface OOClientEventRpc {
    type: "rpc";
    responseKey: number;
    data: Record<string, unknown>;
}

/**
 * Editics addition (no vanilla OO counterpart): a save has been persisted
 * up to the given change index / new version.
 */
export interface OOClientEventSaveDone {
    type: "saveDone";
    savedUpToIndex: number;
    newVersion: number;
}

export type OOClientEvent =
    | OOClientEventAuth
    | OOClientEventMessage
    | OOClientEventCursor
    | OOClientEventGetLock
    | OOClientEventIsSaveLock
    | OOClientEventSaveChanges
    | OOClientEventUnSaveLock
    | OOClientEventUnLockDocument
    | OOClientEventClose
    | OOClientEventAuthChangesAck
    | OOClientEventGetMessages
    | OOClientEventOpenDocument
    | OOClientEventClientLog
    | OOClientEventExtendSession
    | OOClientEventForceSaveStart
    | OOClientEventRpc
    | OOClientEventSaveDone;

// --- Server events (server -> editor) -------------------------------------

/** Positive reply to the editor `auth` handshake. */
export interface OOServerEventAuth {
    type: "auth";
    /** 1 = success. */
    result: number;
    sessionId: string;
    sessionTimeConnect?: number;
    participants: OOParticipantEntry[];
    messages?: OOMessageEntry[];
    /** The fork's own replies send `[]` (no locks at all). */
    locks?: Record<string, OOLockEntry> | unknown[];
    indexUser: number;
    hasForgotten?: boolean;
    jwt?: string;
    g_cAscSpellCheckUrl?: string;
    buildVersion?: string;
    buildNumber?: number;
    licenseType?: number;
    settings?: OOServerAuthSettings;
    openedAt?: number;
    /** Integrator-specific extra fields travel along. */
    [key: string]: unknown;
}

/** The auth lock is held by an established editor; wait for its release. */
export interface OOServerEventWaitAuth {
    type: "waitAuth";
    lockDocument: OOParticipantEntry;
}

/** The participant list changed. */
export interface OOServerEventConnectState {
    type: "connectState";
    participantsTimestamp: number;
    participants: OOParticipantEntry[];
    waitAuth: boolean;
}

/** Changes missed while the editor was away. */
export interface OOServerEventAuthChanges {
    type: "authChanges";
    changes: OOChangeEntry[];
}

/** Chat/comment messages. */
export interface OOServerEventMessage {
    type: "message";
    messages?: Partial<OOMessageEntry>[];
}

/** Cursor positions of the other participants. */
export interface OOServerEventCursor {
    type: "cursor";
    messages: {
        cursor: string;
        time: number;
        user: string;
        useridoriginal: string;
    }[];
}

/** Element-lock grants (reply to the editor's `getLock` request). */
export interface OOServerEventGetLock {
    type: "getLock";
    locks: Record<string, OOLockEntry>;
}

/** Element-locks released. */
export interface OOServerEventReleaseLock {
    type: "releaseLock";
    locks: OOSavedLockEntry[];
}

/** A chunk of persisted changes (drives the editor's save state machine). */
export interface OOServerEventSaveChanges {
    type: "saveChanges";
    changes: OOChangeEntry[] | null;
    changesIndex: number;
    syncChangesIndex: number;
    endSaveChanges: boolean;
    locks?: OOSavedLockEntry[];
    excelAdditionalInfo?: string;
}

/** Ack of an intermediate `saveChanges` chunk. */
export interface OOServerEventSavePartChanges {
    type: "savePartChanges";
    changesIndex: number;
    syncChangesIndex: number;
}

/** Save-lock reply (`false` grants the lock). */
export interface OOServerEventSaveLock {
    type: "saveLock";
    saveLock: boolean;
}

/** Save finished (carries the new save point). */
export interface OOServerEventUnSaveLock {
    type: "unSaveLock";
    index: number;
    time: number;
    syncChangesIndex: number;
}

/** Hard disconnect. */
export interface OOServerEventDrop {
    type: "drop";
    code: number;
    description: string;
}

/** Recoverable protocol error. */
export interface OOServerEventWarning {
    type: "warning";
    code: number;
    message: string;
}

/**
 * License injection, read by the editor's `DocsCoApi._onLicense` — without it
 * the editor idles in WaitAuth and never opens. (Missing from protocol.js's
 * `OOServerEvent` list by mistake.)
 */
export interface OOServerEventLicense {
    type: "license";
    license: {
        /** e.g. 3 = open source. */
        type: number;
        mode: number;
        rights: number;
        buildVersion: string;
        buildNumber: number;
    };
}

export type OOServerEvent =
    | OOServerEventAuth
    | OOServerEventWaitAuth
    | OOServerEventConnectState
    | OOServerEventAuthChanges
    | OOServerEventMessage
    | OOServerEventCursor
    | OOServerEventGetLock
    | OOServerEventReleaseLock
    | OOServerEventSaveChanges
    | OOServerEventSavePartChanges
    | OOServerEventSaveLock
    | OOServerEventUnSaveLock
    | OOServerEventDrop
    | OOServerEventWarning
    | OOServerEventLicense;

// ---------------------------------------------------------------------------
// Editor configuration
// ---------------------------------------------------------------------------
// The vanilla OnlyOffice editor config (see https://api.onlyoffice.com/editors/config/)
// plus the fork's additions.
// Any field not declared below passes through to the original editor unchanged.

/** Document permissions (see https://api.onlyoffice.com/editors/config/document/permissions). */
export interface DocEditorPermissions {
    edit?: boolean;
    download?: boolean;
    print?: boolean;
    copy?: boolean;
    comment?: boolean;
    review?: boolean;
    fillForms?: boolean;
    modifyFilter?: boolean;
    modifyContentControl?: boolean;
    chat?: boolean;
    [key: string]: unknown;
}

/** The `document` section of the editor config. */
export interface DocEditorConfigDocument {
    /** Document id (unique per content; unused in offline mode). */
    key?: string;
    title?: string;
    /**
     * Placeholder URL: never fetched in offline mode, but the vanilla
     * `_checkConfigParams` rejects an empty one.
     */
    url?: string;
    permissions?: DocEditorPermissions;
    [key: string]: unknown;
}

/** The `customization` section of `editorConfig`. */
export interface DocEditorCustomization {
    compactHeader?: boolean;
    /** e.g. "theme-classic-light" / "theme-dark". */
    uiTheme?: string;
    chat?: boolean;
    comments?: boolean;
    help?: boolean;
    about?: boolean;
    feedback?: boolean;
    anonymous?: { request: boolean };
    [key: string]: unknown;
}

/** The `editorConfig` section of the editor config. */
export interface DocEditorConfigEditorConfig {
    mode?: "edit" | "view";
    lang?: string;
    user?: { id?: string; name?: string };
    customization?: DocEditorCustomization;
    [key: string]: unknown;
}

export interface DocEditorEvents {
    onAppReady?: () => void;
    onDocumentReady?: () => void;
    /**
     * Called inside the wrapped `api.asc_Save` with the serialized
     * `Editor.bin` bytes (UTF-8 of the
     * "<signature>;v<version>;<length>;<base64>" wire string); a rejected
     * promise fails the save.
     */
    onSave?: (data: Uint8Array) => void | Promise<void>;
    onError?: (event: { data?: unknown }) => void;
    [key: string]: unknown;
}

/** The editor config accepted by `docEditor.init()`. */
export interface DocEditorConfig {
    /** Fork addition: local mode, no DocumentServer (see `MockServer`). */
    offline?: boolean;
    /**
     * Fork addition: autosave gap in seconds; 0 disables the SDK's periodic
     * autosave (the value is only tested against zero).
     */
    autosave?: number;
    width?: string | number;
    height?: string | number;
    documentType?: "word" | "cell" | "slide" | "pdf";
    document?: DocEditorConfigDocument;
    editorConfig?: DocEditorConfigEditorConfig;
    events?: DocEditorEvents;
    [key: string]: unknown;
}

// ---------------------------------------------------------------------------
// The editor API object living in the iframe
// ---------------------------------------------------------------------------

/**
 * The editor api instance living on the editor iframe window, returned by
 * `docEditor.getApi()`.
 * The surface is huge and build-dependent, so only what this wrapper and its
 * hosts rely on is declared.
 */
export interface OOApi {
    isDocumentModified: () => boolean;
    asc_Save: (...args: unknown[]) => boolean;
    /** Autosave gap in seconds (only used as an on/off flag in practice). */
    asc_setAutoSaveGap: (gap: number) => void;
    /**
     * Serialize the document to the `Editor.bin` wire string (not available
     * in the pdf editor build).
     */
    asc_nativeGetFile: () => string | undefined;
    attachEvent: (name: string, callback: (...args: unknown[]) => void) => void;
    [key: string]: unknown;
}

// ---------------------------------------------------------------------------
// Host page globals
// ---------------------------------------------------------------------------

/**
 * An RGBA color, as returned by `OOHostHooks.getUserColor`
 * (CryptPad's `CPColor` / the editor's `CColor`).
 */
export interface OOHostColor {
    /** 0-255. */
    r: number;
    /** 0-255. */
    g: number;
    /** 0-255. */
    b: number;
    /** 0-255; used by the editor's `CColor` constructor. */
    a: number;
}

/**
 * The descriptor handed to `OOHostHooks.printPdf`/`downloadAs` by the
 * editor's `apiBase.sync_StartAction`-adjacent download pipeline.
 */
export interface OOHostDataContainer {
    /** The in-browser-serialized document bytes (e.g. the PDF to print). */
    data: unknown;
    [key: string]: unknown;
}

/**
 * Cryptpad changes on OnlyOffice often rely on a `window.APP` object to
 * store its hooks. However this `window.APP` was never defined in this
 * codebase (and instead was set in CryptPad code :/).
 * see:
 * - sdkjs/slide/api.js:7320 (hook on editor theme change)
 * - https://github.com/cryptpad/cryptpad/blob/9808cf25c1091d6cf532df13bf5a70ba332f8d4d/www/common/onlyoffice/inner.js#L65)
 *
 * Every hook is optional: each sdkjs call site feature-tests
 * `window.parent.APP.<hook>` and falls back to the vanilla OnlyOffice
 * behaviour (or no-ops) when it is absent.
 */
export interface OOHostHooks {
    /**
     * Print bridge, called by `apiBase` instead of POSTing the PDF to the
     * converter (sdkjs/common/apiBase.js:2658). Only consulted for PDF
     * prints (`options.isPdfPrint`); the callback should be called with
     * `null` (ends the editor's blocking action cleanly) or a server-like
     * `{ status: 'ok', data: url, ... }` object which is then fed back into
     * the editor's print pipeline.
     */
    printPdf?: (
        dataContainer: OOHostDataContainer,
        callback: (result: { status: string; data?: string } | null) => void,
    ) => void;

    /**
     * Download-as bridge, called by `apiBase` instead of POSTing to
     * `/downloadas` (sdkjs/common/apiBase.js:2669). Same callback contract as
     * `printPdf` (`null` to end the action, or the server-like result whose
     * URL is downloaded).
     */
    downloadAs?: (
        dataContainer: OOHostDataContainer,
        callback: (result: { status: string; data?: string } | null) => void,
    ) => void;

    /**
     * Resolve a document image source to a fetchable URL. Called by the
     * image loader when the src looks like a CryptPad media tag
     * ("#channel=...", sdkjs/common/GlobalLoaders.js:639); the callback must
     * receive the blob URL to load, or `""` to fall back to the original src.
     */
    getImageURL?: (name: string, callback: (url: string) => void) => void;

    /**
     * Image-insertion dialog bridge (sdkjs/common/apiBase.js:2804): replaces
     * the editor's upload dialog. The success callback receives an image
     * descriptor whose `url` is added to the document; the error callback is
     * invoked on cancel/failure (raising a non-critical `asc_onError`).
     */
    AddImage?: (
        callback: (res: { url: string; [key: string]: unknown }) => void,
        errorCallback: () => void,
    ) => void;

    /**
     * Image-file upload bridge (sdkjs/common/editorscommon.js:2496): replaces
     * the POST to the converter's upload endpoint. The callback receives an
     * error id (`0`/`c_oAscError.ID.No` on success) and the resulting image
     * urls.
     */
    UploadImageFiles?: (
        files: File[],
        documentId: string,
        documentUserId: string,
        jwt: string,
        callback: (err: number | null, urls: string[]) => void,
    ) => void;

    /**
     * Per-user color override for co-authoring cursors/highlights
     * (sdkjs/common/editorscommon.js:11133). Return the host color for
     * `userId`, or a falsy value to let the editor pick its own deterministic
     * color. Exceptions from the hook are swallowed by the editor.
     */
    getUserColor?: (userId: string) => OOHostColor | null | undefined;

    /**
     * Slides only: the user changed the presentation theme locally
     * (sdkjs/slide/api.js:7320); the host typically persists the theme on the
     * pad and replicates it to co-authors as a `cp_theme` change.
     * Not called for remote/selected-slide changes.
     */
    changeTheme?: (indexTheme: number) => void;

    /**
     * Slides only: a `cp_theme` change arrived from a co-author
     * (sdkjs/common/CollaborativeEditingBase.js:69); the host must re-apply
     * the (remotely updated) theme before the editor applies it itself.
     */
    remoteTheme?: () => void;

    /**
     * Spreadsheets only: the fast-collaborative mode was toggled via
     * `asc_SetFastCollaborative` (sdkjs/cell/api.js:1779); lets the host
     * track/replicate the setting.
     */
    onFastChange?: (bFast: boolean) => void;

    /**
     * Url query args (as a string) appended by the editor to popup urls: the
     * slide reporter window (`index.reporter.html`, sdkjs/slide/api.js:7612)
     * and the index.html links of the outer api script
     * (web-apps/apps/api/documents/api.js:1192). Carries e.g. the pad app
     * route args of the host.
     */
    urlArgs?: string;

    /**
     * Open a hyperlink from the embed viewer (web-apps/apps/common/embed/lib/
     * util/utils.js:134); replaces the viewer's own `window.open` (typically
     * routed by the host to its own navigation). Only read in embed builds.
     */
    openURL?: (url: string) => void;

    [key: string]: unknown;
}

/** Globals installed on the host page window (by the wrapper or host page). */
declare global {
    interface Window {
        /** The patched OnlyOffice API entry point (see `loadAndPatchOOOrig`). */
        DocsAPI: { DocEditor: new (placeholderId: string) => DocEditor };
        /** Hooks introduced by Cryptpad's fork */
        APP?: OOHostHooks;
    }
}

// ---------------------------------------------------------------------------
// End of the typing stuff
// ---------------------------------------------------------------------------

/**
 * How to start the `DocEditor`:
 * ```
 * // 1. Create & initialize the editor
 * const docEditor = new window.DocsAPI.DocEditor("placeholder");
 * const mockServer = ...;
 * window.APP = { getImageURL: ..., AddImage: ... , ...}
 * await docEditor.init(config, mockServer);
 * // 2. Load the document and start the actual edition
 * docEditor.sendMessageToOO(
 *     {
 *         type: 'license',
 *         // ...
 *     }
 * );
 * docEditor.loadBinary(<document binary content>);
 * ```
 */
export class DocEditor implements DocEditorInterface {
    public waitForAppReady: Promise<void>;
    public waitForDocumentReady: Promise<void>;
    private origEditor?: OrigDocEditorInterface;
    private fromOOHandler: EventHandler<FromOO> = new EventHandler();
    private toOOHandler: EventHandler<ToOO> = new EventHandler();
    private placeholderId: string;
    private server: MockServer;
    private fromOOHandle?: HandlerHandle<FromOO>;
    private corruptionWarningHandler: EventHandler<string> = new EventHandler();
    private offline = false;

    constructor(placeholderId: string) {
        this.placeholderId = placeholderId;

        // Cryptpad changes on OnlyOffice often rely on a `window.APP` object to
        // store its hooks. However this `window.APP` was never defined in this
        // codebase (and instead was set in CryptPad code :/).
        // see:
        // - sdkjs/slide/api.js:7320 (hook on editor theme change)
        // - https://github.com/cryptpad/cryptpad/blob/9808cf25c1091d6cf532df13bf5a70ba332f8d4d/www/common/onlyoffice/inner.js#L65)
        const w = window as any;
        w.APP = w.APP ?? {};
    }

    /**
     * Insert the editor iframe and connect the co-authoring "server".
     * Resolves once the editor app is ready. After that the caller is responsible
     * for pushing the `license` and opening the document with `loadBinary()`.
     */
    async init(config: DocEditorConfig, server: MockServer): Promise<void> {
        await scriptLoadedPromise;
        let onAppReady;
        let onDocumentReady;

        this.offline = !!config?.offline;

        this.waitForAppReady = new Promise((resolve) => {
            onAppReady = resolve;
        });

        this.waitForAppReady
            .then(config?.events?.onAppReady ?? noop)
            .catch(noop);

        this.waitForDocumentReady = new Promise((resolve) => {
            onDocumentReady = resolve;
        });

        this.waitForDocumentReady
            .then(() => {
                // Wrapper-level wiring once the document is ready (the editor
                // api exists at this point)
                this.installSaveHook(config);
                this.applyAutosave(config);
            })
            .then(config?.events?.onDocumentReady ?? noop)
            .catch(noop);

        const newConfig = deepAssign(config, {
            events: {
                onAppReady,
                onDocumentReady,
                cryptPadSendMessageFromOO: (msg: { data: { msg: FromOO } }) => {
                    this.fromOOHandler.fire(msg.data.msg);
                },
                cryptPadCorruptionWarningHandler: (msg: {
                    data: { duplicateId: string };
                }) => {
                    this.corruptionWarningHandler.fire(msg.data.duplicateId);
                },
            },
        });

        this.origEditor = new DocEditorOrig(this.placeholderId, newConfig);
        this.toOOHandler.addHandler((msg) =>
            this.origEditor.cryptPadMessageToOO(msg),
        );

        if (this.offline) {
            // Flag the editor iframe as offline for `apiBase.asc_isOffline()`
            // (see sdkjs/common/apiBase.js: 3275)
            const iframe = this.getIframe();
            if (iframe?.contentWindow) {
                (iframe.contentWindow as any).__OOCryptPadOffline = true;
            }
        }

        // Do not resolve before the editor app is ready (see the doc comment
        // on `init` for what that guarantees to the caller).
        await this.waitForAppReady;

        // Connect the co-authoring mock server. Nothing flows over the bridge
        // before the host pushes the license (`sendMessageToOO`), so no
        // message is missed by connecting here.

        this.server = server;

        this.fromOOHandle = this.fromOOHandler.addHandler((msg) => {
            this.server.onMessage(msg);
        });

        if (server.onCorruptionWarning) {
            this.corruptionWarningHandler.addHandler(
                server.onCorruptionWarning,
            );
        }
    }

    destroyEditor() {
        this.fromOOHandle?.remove();
        this.origEditor.destroyEditor();
    }

    getIframe(): HTMLIFrameElement {
        return document.querySelector('iframe[name="frameEditor"]');
    }

    injectCSS(css: string) {
        const head = this.getIframe().contentDocument.querySelector("head");
        const style = document.createElement("style");
        style.innerText = css;
        head.appendChild(style);
    }

    sendMessageToOO(msg: ToOO) {
        this.toOOHandler.fire(msg);
    }

    /**
     * The editor API instance living on the iframe window. Each editor
     * build exposes it its own way: word and slide as `window.editor`, the
     * cell build as `window.editorCell` (its `window.editor` can hold junk,
     * see sdkjs/cell/view/DrawingObjectsController.js:46), pdf not at all.
     * Pick the first candidate that actually is an api (i.e. has `asc_Save`).
     */
    getApi(): OOApi | null {
        const w = this.getIframe()?.contentWindow as any;
        for (const c of [w?.editor, w?.editorCell, w?.Asc?.editor]) {
            if (c && typeof c.asc_Save === "function") {
                return c;
            }
        }
        return null;
    }

    private installSaveHook(config: DocEditorConfig) {
        const saveHandler = config?.events?.onSave;
        if (typeof saveHandler !== "function") return;
        const api = this.getApi();
        if (!api || api.__cpSaveHooked) return;
        api.__cpSaveHooked = true;
        const origAscSave = api.asc_Save;
        api.asc_Save = (...args: any[]) => {
            Promise.resolve()
                .then(() => saveHandler(this.serializeBinary()))
                .then(() => origAscSave.apply(api, args))
                .catch((e) => console.error("DocEditor save failed:", e));
            return true;
        };
    }

    /**
     * Apply `config.autosave` (seconds, 0 = no autosave) via
     * `api.asc_setAutoSaveGap`. The SDK only tests the value against zero —
     * the real periods are hardcoded in `apiBase._autoSave` (immediately
     * while alone with the document lock, otherwise 2 s after the last
     * change when single user, 10 min while co-editing).
     */
    private applyAutosave(config: DocEditorConfig) {
        const gap = Number(config?.autosave);
        if (!Number.isFinite(gap)) return;
        const api = this.getApi();
        if (api && typeof api.asc_setAutoSaveGap === "function") {
            api.asc_setAutoSaveGap(gap);
        }
    }

    /**
     * Serialize the current document to `Editor.bin` bytes fully in-browser
     * via the editor api's `asc_nativeGetFile()` (each editor build writes
     * its own document model). Returns the UTF-8 bytes of the
     * "<signature>;v<version>;<length>;<base64>" wire string
     * ("DOCY"/"XLSY"/"PPTY" per build) — the exact form `loadBinary`
     * accepts. Not available in the pdf editor build (it edits real PDF
     * bytes).
     */
    private serializeBinary(): Uint8Array {
        const api = this.getApi();
        const native = api?.asc_nativeGetFile?.();
        if (!native) {
            throw new Error("could not serialise the document to Editor.bin");
        }
        return new TextEncoder().encode(native);
    }

    serviceCommand(command: string, data: any) {
        this.origEditor.serviceCommand(command, data);
    }

    showMessage(...args: any[]) {
        return this.origEditor.showMessage(...args);
    }
    processSaveResult(...args: any[]) {
        return this.origEditor.processSaveResult(...args);
    }
    processRightsChange(...args: any[]) {
        return this.origEditor.processRightsChange(...args);
    }
    denyEditingRights(...args: any[]) {
        return this.origEditor.denyEditingRights(...args);
    }
    refreshHistory(...args: any[]) {
        return this.origEditor.refreshHistory(...args);
    }
    setHistoryData(...args: any[]) {
        return this.origEditor.setHistoryData(...args);
    }
    setEmailAddresses(...args: any[]) {
        return this.origEditor.setEmailAddresses(...args);
    }
    setActionLink(...args: any[]) {
        return this.origEditor.setActionLink(...args);
    }
    processMailMerge(...args: any[]) {
        return this.origEditor.processMailMerge(...args);
    }
    downloadAs(...args: any[]) {
        return this.origEditor.downloadAs(...args);
    }
    attachMouseEvents(...args: any[]) {
        return this.origEditor.attachMouseEvents(...args);
    }
    detachMouseEvents(...args: any[]) {
        return this.origEditor.detachMouseEvents(...args);
    }
    setUsers(...args: any[]) {
        return this.origEditor.setUsers(...args);
    }
    showSharingSettings(...args: any[]) {
        return this.origEditor.showSharingSettings(...args);
    }
    setSharingSettings(...args: any[]) {
        return this.origEditor.setSharingSettings(...args);
    }
    insertImage(...args: any[]) {
        return this.origEditor.insertImage(...args);
    }
    setMailMergeRecipients(...args: any[]) {
        return this.origEditor.setMailMergeRecipients(...args);
    }
    setRevisedFile(...args: any[]) {
        return this.origEditor.setRevisedFile(...args);
    }
    setFavorite(...args: any[]) {
        return this.origEditor.setFavorite(...args);
    }
    requestClose(...args: any[]) {
        return this.origEditor.requestClose(...args);
    }
    grabFocus(...args: any[]) {
        return this.origEditor.grabFocus(...args);
    }
    blurFocus(...args: any[]) {
        return this.origEditor.blurFocus(...args);
    }
    setReferenceData(...args: any[]) {
        return this.origEditor.setReferenceData(...args);
    }

    /**
     * Open a document from raw bytes. This function should be called after `init`.
     */
    loadBinary(data: Uint8Array) {
        const iframe = this.getIframe().contentWindow;
        const msg: any = {
            command: "openDocumentFromBinary",
            data: data,
        };
        iframe.postMessage(msg, iframe.location.origin, [data.buffer]);
    }

    /**
     * Trigger a save. This calls the editor api's `asc_Save()` (same code path
     * as for the toolbar Save button).
     */
    save() {
        const api = this.getApi();
        if (api && typeof api.asc_Save === "function") {
            api.asc_Save();
        }
    }
}

type FromOO = OOClientEvent;
type ToOO = OOServerEvent;

export interface DocEditorInterface {
    showMessage(...args: any[]): any;
    processSaveResult(...args: any[]): any;
    processRightsChange(...args: any[]): any;
    denyEditingRights(...args: any[]): any;
    refreshHistory(...args: any[]): any;
    setHistoryData(...args: any[]): any;
    setEmailAddresses(...args: any[]): any;
    setActionLink(...args: any[]): any;
    processMailMerge(...args: any[]): any;
    downloadAs(...args: any[]): any;
    serviceCommand(command: string, data: any): void;
    attachMouseEvents(...args: any[]): any;
    detachMouseEvents(...args: any[]): any;
    destroyEditor(): void;
    setUsers(...args: any[]): any;
    showSharingSettings(...args: any[]): any;
    setSharingSettings(...args: any[]): any;
    insertImage(...args: any[]): any;
    setMailMergeRecipients(...args: any[]): any;
    setRevisedFile(...args: any[]): any;
    setFavorite(...args: any[]): any;
    requestClose(...args: any[]): any;
    grabFocus(...args: any[]): any;
    blurFocus(...args: any[]): any;
    setReferenceData(...args: any[]): any;
    loadBinary(data: Uint8Array): void;
    save(): void;
}

export interface OrigDocEditorInterface extends DocEditorInterface {
    cryptPadMessageToOO(msg: OOServerEvent): void;
}

export interface MockServer {
    onMessage: (msg: OOClientEvent) => void;
    onCorruptionWarning?: (duplicateId: string) => void;
}

async function loadAndPatchOOOrig() {
    let myScriptSrc: string;
    let myScriptElement: HTMLScriptElement;

    // TODO document.currentScript does not return the correct tag?
    // Use this as a workaround:
    for (const e of document.getElementsByTagName("script")) {
        try {
            const pathname = new URL(e.src).pathname;
            if (pathname.endsWith("web-apps/apps/api/documents/api.js")) {
                myScriptSrc = e.src;
                myScriptElement = e;
                break;
            }
        } catch (error) {
            if (error instanceof TypeError) {
                // e.src is not a valid URL -> ignore
            } else {
                throw error;
            }
        }
    }
    const script = document.createElement("script");
    script.setAttribute("type", "text/javascript");
    const newUrl = new URL("api-orig.js", myScriptSrc);
    newUrl.search = new URL(myScriptSrc).search;
    script.setAttribute("src", newUrl.href);
    const scriptLoadedPromise = waitForEvent(script, "load");

    myScriptElement.after(script);

    // Setup window.DocsAPI.DocEditor
    const w = window as any;
    w.DocsAPI = w.DocsAPI ?? {};
    w.DocsAPI.DocEditor = DocEditor;

    await scriptLoadedPromise;
    // Setup window.DocsAPI.DocEditor again after the original editor replaced it
    DocEditorOrig = w.DocsAPI.DocEditor;
    w.DocsAPI.DocEditor = DocEditor;
}

const scriptLoadedPromise = loadAndPatchOOOrig();
