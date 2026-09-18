import { EventHandler, HandlerHandle } from "./eventHandler";
import { deepAssign, noop, waitForEvent } from "./utils";

let DocEditorOrig: any;

/**
 * How to start the `DocEditor`:
 * ```
 * const docEditor = new DocsAPI.DocEditor("placeholder");
 * const mockServer = ...;
 * await docEditor.init(config, mockServer);
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
    }

    /**
     * Insert the editor iframe and connect the co-authoring "server".
     * Resolves once the editor app is ready. After that the caller is responsible
     * for pushing the `license` and opening the document with `loadBinary()`.
     */
    async init(config: any, server: MockServer): Promise<void> {
        await scriptLoadedPromise;
        let onAppReady;
        let onDocumentReady;

        this.offline = !!config?.offline;

        this.installHostHooks(config);

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
        this.connectMockServer(server);
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

    private installHostHooks(config: any) {
        const w = window as any;
        const events = config?.events ?? {};
        const endAction = (dataContainer: any, cb: (obj?: any) => void) =>
            cb(null);
        if (events.onPrintPdf || this.offline) {
            w.APP.printPdf = events.onPrintPdf ?? endAction;
        }
        if (events.onDownloadAs || this.offline) {
            w.APP.downloadAs = events.onDownloadAs ?? endAction;
        }
    }

    private connectMockServer(server: MockServer) {
        this.server = server;

        const w = window as any;
        w.APP.getImageURL = server.getImageURL
            ? (name: string, callback: (url: string) => void) => {
                  server
                      .getImageURL(name)
                      .then(callback)
                      .catch((e) => console.error(e));
              }
            : (name: string, callback: (url: string) => void) => callback("");

        this.fromOOHandle = this.fromOOHandler.addHandler((msg) => {
            this.server.onMessage(msg);
        });

        if (server.onCorruptionWarning) {
            this.corruptionWarningHandler.addHandler(
                server.onCorruptionWarning,
            );
        }
    }

    /**
     * The editor api instance living on the iframe window.
     */
    private getApi(): any {
        const w = this.getIframe()?.contentWindow as any;
        return w?.editor ?? w?.Asc?.editor ?? null;
    }

    private installSaveHook(config: any) {
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
     * `api.asc_setAutoSaveGap`. NB: the SDK uses the value only as an on/off
     * flag — `apiBase._autoSave` tests `0 !== api.autoSaveGap` (pdf's
     * `_autoSave` override adds `|| Is_Fast()`); it is never used as a
     * duration. The real periods are hardcoded: `onDocumentContentReady`
     * unconditionally starts a 40 ms `_autoSave` ticker which either
     *
     *   - saves immediately on every tick while a single user holds the
     *     document lock (`canUnlockDocument`: the co-authoring protocol's
     *     solo state, set by `_unlockDocument()` — no interval consulted at
     *     all), or
     *   - runs `_autoSaveInner` (only when the flag is on), which saves
     *     `autoSaveGapFast` (2 s) after the last change for a single user
     *     (`CollaborativeEditing.m_nUseType === 1`) and `autoSaveGapSlow`
     *     (10 min) while co-editing (`m_nUseType <= 0`); a change younger
     *     than `intervalWaitAutoSave` (1 s) postpones it, and in fast
     *     co-editing it exchanges changes instead of saving
     *     (`Continue_FastCollaborativeEditing`).
     *
     * `autosave: 0` disables the periodic autosave entirely while
     * `canUnlockDocument` is false — as in this wrapper's local mode, where
     * the editor is single-user from the start so nothing ever calls
     * `_unlockDocument()`.
     */
    private applyAutosave(config: any) {
        const gap = Number(config?.autosave);
        if (!Number.isFinite(gap)) return;
        const api = this.getApi();
        if (api && typeof api.asc_setAutoSaveGap === "function") {
            api.asc_setAutoSaveGap(gap);
        }
    }

    /**
     * Serialise the current document to raw OnlyOffice `Editor.bin` bytes
     * (`.bin`, "DOCY…") fully in-browser — the same serializer the api's
     * `asc_nativeGetFile()` uses (`new BinaryFileWriter(doc).Write()`), with
     * no server and no `/downloadas` round-trip. This mirrors Cryptpad's
     * integration-mode `getContent()`. Word uses `AscCommonWord`, cell
     * `AscCommonExcel`; other editors fall back to `asc_nativeGetFile()`
     * (whose string form is decoded by `decodeNativeFileString`).
     */
    private serializeBinary(): Uint8Array {
        const w = this.getIframe()?.contentWindow as any;
        const api = this.getApi();
        const doc = api?.WordControl?.m_oLogicDocument;
        if (!doc) throw new Error("logic document not ready");
        const BFW =
            w?.AscCommonWord?.BinaryFileWriter ??
            w?.AscCommonExcel?.BinaryFileWriter ??
            w?.AscCommon?.BinaryFileWriter;
        if (typeof BFW === "function") {
            // Write(true) returns the raw bytes; Write() (no arg) returns a
            // "DOCY;v<version>;<length>;<base64>" string.
            const out = new BFW(doc).Write(true);
            if (out && typeof out.byteLength === "number") return out;
            if (typeof out === "string") return decodeNativeFileString(out);
        }
        const native = api?.asc_nativeGetFile?.();
        if (native && typeof native.byteLength === "number") return native;
        if (typeof native === "string") return decodeNativeFileString(native);
        throw new Error("could not serialise the document to Editor.bin");
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

type FromOO = any;
type ToOO = any;

interface DocEditorInterface {
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

interface OrigDocEditorInterface extends DocEditorInterface {
    cryptPadMessageToOO(msg: ToOO): void;
}

interface MockServer {
    getImageURL?: (name: string) => Promise<string>;
    onMessage: (msg: FromOO) => void;
    onCorruptionWarning?: (duplicateId: string) => void;
}

/**
 * Decode the `BinaryFileWriter.Write()` string form (i.e.
 * `"DOCY;v<version>;<byteLength>;<base64>"`) into raw `Editor.bin` bytes.
 * (Used as a fallback by `DocEditor.serializeBinary` for editors whose
 * `BinaryFileWriter` namespace the wrapper does not know about.)
 */
function decodeNativeFileString(native: string): Uint8Array {
    const parts = native.split(";");
    const b64 = parts[parts.length - 1];
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
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
