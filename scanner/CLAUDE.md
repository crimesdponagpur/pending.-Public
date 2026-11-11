# ScanOS: notes for any Claude working in `scanner/`

ScanOS is a standalone mobile scanner PWA (iPhone + Android), hosted on GitHub Pages at
`https://crimesdponagpur.github.io/pending.-Public/scanner/`. The owner (ND) writes Marathi/Hindi in Roman script.
**Reply to ND in simple Marathi (Devanagari). All UI text in the app is Marathi.**

## Layout
- `index.html`: the whole app (HTML + CSS + one IIFE of vanilla JS, ~4000 lines). No build step, no framework.
- `sw.js`: service worker. Bump the cache name (`scanos-vN`) on every release, or users keep the old version. Also handles the share-target POST (`?share=1`).
- `manifest.webmanifest`, `config.js`, icons.
- The Apps Script backend (`Code.gs`) is **not in this repo** (it holds the owner's setup). The app talks to it with `api(fn, args)`; functions: `info, ocr, translate, ai, image, save`. AI features use `api('ai', {task:'chat', prompt, images, json:true})`. Server changes need ND to paste the new Code.gs and Deploy a New version.

## How the code is organised (search the `/* ---------- ... ---------- */` headings in index.html)
- IndexedDB: store `docs` (doc metadata incl. `pages[]`), store `blobs` with keys `<pageId>:o` (original), `:o0` (before AI), `:p` (processed JPEG), `:t` (thumbnail), `<docId>:pdf`, `office:lh`, `office:sig<N>`.
- Pages are rendered from the original with `renderPage` (warp quad, rotate, filter, marks). Marks use coordinates normalised 0..1.
- Screens are `.view` sections, navigated with `go()/back()` and `setView()`; dialogs use `openDlg(html, bind)`; busy spinner `busy()/done()`; errors `fail(e)`.
- Doc actions are buttons with `data-a="x"` mapped in the `ACTS` object. New AI doc features usually use `aiDocAction(...)` and `tableDlg(...)`.
- Everything works offline except features that call the server.

## Rules
1. Work on a new branch from the latest `main`, open a PR, and do not merge unless ND asks. Only one Claude should work at a time: a single big file means merge conflicts.
2. Keep AI output honest: never invent names, numbers or dates. Health, medicine and legal features must keep the "not a diagnosis / advice" notice.
3. Keep touch response fast (`touch-action: manipulation`), keep Marathi labels short, test at 430px width.
4. Never commit secrets (server URL key, access code, API keys). Those live in the owner's phone / Script Properties.
5. Bump `sw.js` cache version in every PR that changes `index.html`.

## Testing without a phone
Serve the repo root (`python3 -m http.server 8771`) and drive `http://localhost:8771/scanner/` with Playwright at 430x900. Set `localStorage` keys `sm-onb=1`, `sm-api`, `sm-key`, `sm-ai=1` and mock the server with `route()` (the app POSTs JSON `{key, fn, a}` as text/plain and expects `{ok:true, d:...}`). Real-device checks (camera, Face ID, share target) can only be done by ND.
