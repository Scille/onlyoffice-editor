# OnlyOffice offline mode demo

A small web app demonstrating the offline (no-server, fully client-side)
document-editing mode of the CryptPad OnlyOffice wrapper.

Everything happens in the browser: there is no DocumentServer, no co-authoring
network, no `/downloadas` converter. The document is stored as an OnlyOffice
`Editor.bin` (`.bin`) in `localStorage` and (re)opened with
`DocEditor.loadBinary()`.

Pick the editor type with a URL parameter: `/?type=word|cell|slide` (default
`word`). On first run the document is seeded with the empty-document
`Editor.bin` template the demo server extracts from the sdkjs sources
(`/empty.bin?type=…`) — each editor build has its own `Editor.bin` signature
(`DOCY`/`XLSY`/`PPTY`), so stored documents are keyed per type in
`localStorage`.

## Running the demo

```sh
cd demo
./build.sh           # Build sdkjs/web-app/onlyoffice-editor (done once)
npm install          # installs playwright (for the smoke test)
npm run demo         # starts the server on http://localhost:4013/
# open http://localhost:4013/?type=word  (or ?type=cell / ?type=slide)
```

The demo server serves:
- the demo's own files (`index.html`, `app.js`) from the `demo/` directory
- the built OnlyOffice client bundle (JS/CSS/wasm/fonts/dictionaries) from
  `../sdkjs/deploy`

## Running the smoke test

```sh
# Requires playwright browsers. Either set PLAYWRIGHT_BROWSERS_PATH
# or run `npx playwright install` first.

# one editor type (word | cell | slide):
PLAYWRIGHT_BROWSERS_PATH=/path/to/playwright/browsers npm run smoke -- word

# all three editor types in sequence:
PLAYWRIGHT_BROWSERS_PATH=/path/to/playwright/browsers npm run smoke:all
```

Each run verifies the full round-trip for the chosen editor type: boot,
`loadBinary` + auth handshake, TWO type-specific edits, Save to `localStorage`,
reload (both edits survive) and delete (storage reseeded from the empty
template).

The two-edit part doubles as the block-lock regression test: the editors
request element locks even in single-user mode, and `LocalMockServer` must
grant them to the requester's composite `<idOriginal><indexUser>` id (see
`answerGetLock` in `app.js`) — otherwise the editor considers its own blocks
locked by someone else and discards the subsequent edits of the same block
(only one modifiable cell / one typed character per text box). The test
checks the granted lock state (2 = "acquired by me", not 3 = "held by someone
else") and re-runs the same lock gate the editor uses (adding a worksheet /
adding or modifying a shape) to make sure the same kind of edit is still
allowed. Word takes no block locks for plain text edits, so it only checks
that the two text edits stick.
