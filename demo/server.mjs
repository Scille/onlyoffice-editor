// Minimal static file server for the local-mode demo.
//
// Serves two roots:
//   - the demo's own files (index.html / app.js) from this directory
//   - the built OnlyOffice client bundle (JS/CSS/wasm/fonts/dictionaries)
//     from ../sdkjs/deploy
//
// Plus one dynamic endpoint: `/empty.bin?type=word|cell|slide` — the
// empty-document `Editor.bin` seed, extracted from the sdkjs *sources*
// (the deployed sdk-all bundles do not contain `AscCommon.getEmpty`, the
// in-memory template the SDK's own empty-doc path would use — it was never
// compiled in, which is also why the SDK's magic-URL empty-document path
// crashes). Each editor build's template lives as a string literal in its
// source tree; we regex it out and serve the wire string as bytes — the
// same `<signature>;v<version>;<length>;<base64>` form the editor's open
// path requires (the decoded base64 payload alone would be rejected by
// `getEditorByBinSignature`, no signature header) and the same form
// `DocEditor.save()` stores.
//
// Usage:  node server.mjs [PORT]      (npm run demo)
import http from "node:http";
import { readFile, stat } from "node:fs/promises";
import { resolve, extname, normalize, join } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const DEMO = __dirname;                       // demo/ (index.html, app.js)
const DEPLOY = resolve(__dirname, "../sdkjs/deploy"); // built OnlyOffice client
const SDKJS = resolve(__dirname, "../sdkjs");          // sdkjs sources
const PORT = parseInt(process.argv[2] || process.env.PORT || "4013", 10);

// `return "<signature>;v<version>;<length>;<base64>";` in each editor's
// empty-document template module. Signatures: DOCY (word), XLSY (cell),
// PPTY (slide).
const EMPTY_TEMPLATE = {
    word: { file: "word/document/editor.js", signature: "DOCY" },
    cell: { file: "cell/document/empty.js", signature: "XLSY" },
    slide: { file: "slide/document/editor.js", signature: "PPTY" },
};
const emptyCache = new Map();

/** Extract the empty-document template wire string for `type` as bytes. */
async function emptyBin(type) {
    if (emptyCache.has(type)) return emptyCache.get(type);
    const spec = EMPTY_TEMPLATE[type];
    if (!spec) throw new Error("unknown editor type: " + type);
    const src = await readFile(join(SDKJS, spec.file), "utf8");
    const m = src.match(new RegExp('return "(' + spec.signature + '[^"]*)";'));
    if (!m) {
        throw new Error(
            "no " + spec.signature + " template in " + spec.file +
                " — is the sdkjs source tree in sync?",
        );
    }
    // Serve the "<signature>;v<version>;<length>;<base64>" wire string as
    // bytes — the `Editor.bin` form `loadBinary` accepts (see the header
    // comment).
    const bytes = Buffer.from(m[1], "utf8");
    emptyCache.set(type, bytes);
    return bytes;
}

const MIME = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".mjs": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".wasm": "application/wasm",
    ".png": "image/png", ".jpg": "image/jpeg", ".gif": "image/gif",
    ".svg": "image/svg+xml",
    ".ttf": "font/ttf", ".otf": "font/otf",
    ".woff": "font/woff", ".woff2": "font/woff2",
    ".bin": "application/octet-stream",
    ".xml": "application/xml; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
    ".map": "application/json; charset=utf-8",
    ".dic": "application/octet-stream", ".aff": "application/octet-stream",
};

// The binary Editor.bin / AllFonts.js embed font paths like "/fonts/arial.ttf"
// while the editor also prepends its font base "/fonts/", producing doubled
// requests "/fonts//fonts/arial.ttf". Collapse the double slash and the
// redundant "fonts/fonts" segment so the font files in <DEPLOY>/fonts/ resolve.
function collapseFontPath(urlPath) {
    let collapsed = urlPath;
    while (collapsed.indexOf("//") !== -1) collapsed = collapsed.split("//").join("/");
    collapsed = collapsed.split("/fonts/fonts/").join("/fonts/");
    return collapsed;
}

// Resolve `urlPath` against `root`, guarding against path traversal, returning
// the file stat if it resolves to a regular file (following index.html for
// directories), else null.
async function tryRoot(root, urlPath, collapseFonts) {
    let p = urlPath;
    if (collapseFonts) p = collapseFontPath(p);
    let path = normalize(join(root, p));
    if (!path.startsWith(root)) return null;
    let s = await stat(path).catch(() => null);
    if (s && s.isDirectory()) {
        path = join(path, "index.html");
        s = await stat(path).catch(() => null);
    }
    if (!s || !s.isFile()) return null;
    return path;
}

http.createServer(async (req, res) => {
    try {
        const u = new URL(req.url, `http://localhost:${PORT}`);
        const url = decodeURIComponent(u.pathname);
        // Empty-document seed: /empty.bin?type=word|cell|slide (default word).
        if (url === "/empty.bin") {
            const type = (u.searchParams.get("type") || "word").toLowerCase();
            try {
                const bytes = await emptyBin(type);
                res.writeHead(200, {
                    "Content-Type": "application/octet-stream",
                    "Content-Length": bytes.length,
                    "Cross-Origin-Resource-Policy": "cross-origin",
                });
                res.end(bytes);
            } catch (e) {
                res.writeHead(404); res.end(String(e));
            }
            return;
        }
        // Demo's own files first (index.html, app.js), then the built bundle.
        const path =
            (await tryRoot(DEMO, url, false)) ||
            (await tryRoot(DEPLOY, url, true));
        if (!path) {
            res.writeHead(404); res.end("not found: " + url); return;
        }
        const data = await readFile(path);
        res.writeHead(200, {
            "Content-Type": MIME[extname(path).toLowerCase()] || "application/octet-stream",
            "Content-Length": data.length,
            "Cross-Origin-Resource-Policy": "cross-origin",
        });
        res.end(data);
    } catch (e) {
        console.error("ERR", req.url, e);
        res.writeHead(500); res.end("error");
    }
}).listen(PORT, () => {
    console.log(`demo at http://localhost:${PORT}/  (demo ${DEMO}, deploy ${DEPLOY})`);
});
