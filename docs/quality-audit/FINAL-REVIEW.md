# FINAL-REVIEW — office-quality-all @9ffd696 (adversarial, read-only)

## Critical

**C1. COUNTIF wildcard → catastrophic regex backtracking that blocks the whole server** — `src/collab/condFormatShared.ts:1124` (xlsx lane, 0adb412)
- Each `*` in a COUNTIF criterion becomes `.*` in `new RegExp('^…$','i')`. That pattern is then run against every distinct string in the range.
- Measured: `^.*a.*a.*a.*a.*a.*a.*b$` against a 100-char string of "a" took **25.7 s**. With 50 chars it took 1 s. Eight wildcards against 200 chars did not finish within the timeout.
- Reachable from two places, both on the Node main thread:
  1. `POST /api/view/xlsx-cf-ranges`. Any edit-token holder can send `crit` strings up to 300 chars, up to 2,000 of them (`src/xlsx/cfRangeValues.ts:312`). The path is `answerCfRemote` → `cfRangeAnswer` → `countIf`. A cell holding ~100+ repeated letters is enough.
  2. The server-side viewer bake (`src/xlsx/viewerGrid.ts:160` `bakeCondFormats` → `paintCell` → `CF_RANGE_FUNCS.COUNTIF` with local tables). It runs on upload (`runGridExtraction`) and on stale-grid regeneration, including the new daily `cfDay` regeneration. So a crafted xlsx uploaded by anyone hangs the server for everyone: a rule like `COUNTIF($A:$A,"*a*a*a*a*a*a*a*b")>0` plus one long cell. The bake's time budget cannot interrupt a single `re.test`. (Upload reachability ~85% confidence: path traced, not run end to end. The route path is direct.)
- Fix: drop the regex and use a linear wildcard matcher (greedy two-pointer `*`/`?` match, O(n·m)). Also handle Excel's `~*`/`~?` escapes. Optionally cap the operand to ≤255 chars, as Excel does.

**C2. Quadratic regex on every docx LibreOffice conversion (upload and every save)** — `src/converters/libreOffice.ts:255` (docx lane, `docxViewLayoutFixes`)
- The regex is `/<w:sectPr\b[\s\S]*?<\/w:sectPr>(?![\s\S]*<w:sectPr\b)/`. Each lazy extension re-runs the lookahead to the end of the string.
- Crafted input: k closed `<w:sectPr>…</w:sectPr>`, then a tail that starts with an unclosed `<w:sectPr ` and runs 200 KB. Measured by calling `docxViewLayoutFixes` directly:
  - k=200 / 214 KB: **2.0 s**
  - k=400: **7.1 s**
  - k=800 / 254 KB: **27.5 s**
  - That is ×4 per ×2, so a 1–2 MB document.xml runs for minutes.
- It runs synchronously in the main process (`preprocess` at `libreOffice.ts:371`): at upload conversion, and in `flushToStore` after every docx save.
- The same function's `/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g` (`:260`) is also quadratic on unclosed tags. 89 KB took 55 ms, which extrapolates to about 7 s at 1 MB.
- Fix: find the last sectPr with `lastIndexOf('<w:sectPr')` + `indexOf('</w:sectPr>', at)`. Scan anchors with an indexOf loop that stops when there is no closing tag. Or skip the fix above a size threshold and return the input unchanged.

## Important

**I1. `/api/view/info` now does a full, uncached docx unzip and parse on every request, including read-only viewing** — `src/routes/viewer.ts:721`, `src/docx/import.ts:1686`
- `getDocxRenderExtras` calls `guardedUnzipSync(docxBytes)` with **no filter**, so it inflates every media part too. It then parses document.xml and all header/footer/note parts synchronously.
- Before this branch, the info route only did filtered unzips, and the full-document check (`docxHasUnsupportedContent`) ran only for `canEdit`.
- `public/js/viewer.js:2817` (the PDF viewer, read-only tokens) calls the same route, so every docx open now pays this cost even though only the editor uses `docxRender`.
- Measured: synthetic docx with a 27 MB document.xml took **1.5 s blocking per call**; page setup alone took 60–90 ms. Images add inflate time on top.
- Fix: compute `docxRender` only when the editor asks (e.g. an `editor=1` query, or only for `canEdit` plus the edit page). Unzip with a filter (`word/*.xml`, `_rels`). Cache by docId and file version.

**I2. Close-grace keeps up to 20 rooms of any size; only large xlsx is excluded** — `src/collab/room.ts:3062-3076` (pptx lane mechanism, now applied to every kind)
- PDF rooms hold `originalPdfBytes` and possibly `pdfEncryptedFile` (uploads up to 500 MB). pptx and docx rooms hold the full Y.Doc. Medium xlsx rooms hold lazyGrid and the cf tally.
- Scenario: one user opens and closes 20 large PDFs or decks within 45 s. All 20 stay resident, where before each was freed on close. On the memory-store deployment (4 GB budget) this can push the process toward OOM.
- INTEGRATION.md only weighed the cf tally. Confidence ~75%: the peak depends on real usage.
- Fix: in `lingerRoom`, skip rooms whose original bytes exceed a threshold (like the large-xlsx exclusion), or cap lingering by total bytes, not by count.

## Minor

**M1. 6142bae cleanup can run before an in-flight linger flush finishes** — `src/collab/room.ts:2517-2527` with `:1983`, `:2183`
- Sequence:
  1. The last user leaves, so `closeConn` → linger → `queueFlush` starts.
  2. `flushToStore` sets `editedSinceSave=false` at the top, then awaits `readOriginal`/`saveOriginal`.
  3. In that window `forceCloseRoom`, LRU eviction of the 21st lingering room, or the grace timer (if the flush takes over 45 s) calls `teardownRoom`.
  4. Teardown sees `editedSinceSave=false`, so the `.finally` runs immediately (it does not wait on the serialized flush): it deletes `inheritedFragments`/`writtenFragments` and destroys the ydoc.
  5. The in-flight `saveDocxFragments` then runs with `inherited = new Set()`. It writes last generation's inherited fragments again, breaking "one generation only", and re-adds a `writtenFragments` entry (up to 4 MB) that leaks until that doc is next opened.
- No user data is lost; fragments are kept rather than dropped.
- The same window lets `ydoc.destroy()`, `persistChangeLog` and `closeLazyGridStore` run while a pptx/xlsx flush is still in flight. That is new with the grace period; before it, teardown always enqueued its own serialized flush.
- Fix: in teardown, always chain the cleanup through `runSerialized(docId, …)` (a no-op step when nothing is edited). Then `.finally` runs only after any in-flight flush.

**M2. Request racing teardown on a closed lazy store** — `src/collab/room.ts:640-652`
- `getCfRemoteAnswer` awaits up to 2 s, then reads `room.lazyGrid`. Teardown at grace end can `closeLazyGridStore` in between, and the route returns 500 instead of 409.
- It can also refill a tally that `cfTallyRelease` already removed from `liveTallies`. The tally is GC'd with the room, so it does not leak.
- Fix: after the awaits, re-check `rooms.get(docId) === room` and return null.
- Confidence ~70% that a closed store throws, as opposed to returning empty.

**M3. Possible quadratic scan in docx export clone path** — `src/docx/export.ts:332`, `:354`
- These use tempered `(?:(?!</w:footnote>)[\s\S])*` scans, plus `Math.max(0, ...matches)` (a spread of more than ~100k ids throws RangeError).
- Only reachable when a footnote reference is cloned on save, with a crafted footnotes part.
- Not measured; confidence ~50%.

## Checked and clean

- **Merge 2c0219a, room.ts:** the xlsx-side immediate-close block (cancelPendingUnseeds, rooms/largeRoomDocIds delete, cfTallyRelease, persistChangeLog, closeLazyGridStore, ready→flush→destroy, compared against 516c9e1) is reproduced exactly in `teardownRoom`. Nothing was lost or duplicated. `forceCloseRoom` and LRU eviction both go through teardown.
- **Merge dd4cc0e:** the editor.ts imports (guardRoomEpoch, docxLayout, docxPaste, types) and the viewer.ts room.js import list union are complete. `guardRoomEpoch` is called in all four editors (editor, grid-editor, editor-slides, editor-pdf).
- **run.mjs** resolutions match INTEGRATION.md.
- **6142bae new-room race:** the map deletion is guarded by `!rooms.has(docId)`, and `rooms.set` is synchronous in `getOrCreateRoom`. A new room created before the finally is never wiped. A new room created after it reloads from `docxfrag.json`. No unsaved-data loss found (but see M1).
- **liveTallies:** a tally is created for every room kind and released only in teardown, which is the only `rooms.delete` site. No leak path.
- **Earlier fixes survived the merges:**
  - chart SVG: `esc()` on all text; colors only via `safeChartColor` (#hex); file charts pass `sanitizeGridChart` with anchors clamped to 16,383 cols and 1,048,575 rows.
  - cf-ranges caps: `CF_MAX_*`, 128 KB body, 240/min, canEdit only.
  - docx font CSS: `safeFontName` allow-list in `fontStack`; per-rule `insertRule` in try/catch; selectors via `CSS.escape`.
  - pptx ReDoS: GUID check plus `indexOf` in `loPreprocess.ts`. The remaining dynamic `new RegExp` sites in src interpolate only constants or escaped values, except C1.
- **Routes:**
  - `slide-media` checks that the path prefix matches the document kind (`ppt/` vs `word/`), with the schema pattern `^(ppt|word)/media/[^/]+$`.
  - `docx-drawing` uses a strict ref pattern, an editor-data token, a docx-only check and a heavy rate limit.
  - `xlsx-chart-images` uses a view token with sanitized charts.
  - `grid/meta` re-sanitizes styles; `sanitizeCellStyle` covers every `GridStyle` key, so no view regression from dropped keys.
  - The viewer.js `<style>` sinks (`fc`, `cb`, `ci`, `x`, `xi`) only receive sanitized values.
- **hwp/hwpx save path and legacy download:** untouched. No room-external `saveOriginal` races with lingering rooms (hwp has no room).
- **memoryExpiry:** only delayed 45 s by `isRoomOpen` during grace. Large xlsx rooms do not linger.
- **Removed `#unsupported-warning`** from editor.html: no remaining reference in the docx editor (grid editor keeps its own).

## Not checked

- The full client code of docxLayout, grid-editor and editor-slides beyond the CSS/HTML sinks noted above.
- Whether C1 also runs in the client bake (that would only freeze the editor's own tab).
- `npx tsc` and vitest were not re-run (reported green). No end-to-end server repro of C1/C2; the regex and function were measured directly in isolation.
- cad/dwg paths (no src diff for them). PDF editor behaviour under the grace period beyond reading the flush code.
- The structured-reference rewrite logic in the xlsx table rules (lane-reviewed).

---

## 재검토 (9ffd696..fe23856, read-only)

How I checked: read the diffs. Ran direct timing probes and a 200,000-case differential fuzz (old vs new `docxViewLayoutFixes`). Ran the six new test files: 46/46 pass. The full suite was not re-run.

| Item | Status | Notes |
|---|---|---|
| C1 | **부분 반영** | The exponential case is gone, but one request can still take hours (see below). |
| C2 | 반영 | |
| I1 | 반영 | |
| I2 | 반영 (one gap, Minor) | |
| M1 | 반영 | |
| M2 | 반영 | |
| M3 | 반영 (one sibling left, Minor) | |

### C1 — 부분 반영 (`src/collab/wildcard.ts`, `condFormatShared.ts:1124`)

What is fixed:
- **No more exponential blow-up.** `*a`×8+`*b` against 32,767 chars now takes about 0 ms. `*a`×150 takes 0.5 ms.
- **Excel semantics hold.** 26 hand cases pass:
  - `~*` `~?` `~~` are literals; a lone `~` is a literal.
  - `?` matches exactly one character; `*` matches the empty string.
  - Case is ignored on the CF path.
  - The `Symbol.search` position matches what SEARCH expects.
- **The library override covers every call site in fast-formula-parser 1.0.19:**
  - `helpers.js:360,370`: Criteria `wc`, used by COUNTIF, SUMIF and AVERAGEIF through `.test`.
  - `reference.js:142,355`: MATCH, VLOOKUP and HLOOKUP through `.test`.
  - `text.js:284`: SEARCH through `String.search`.
  - The patch replaces the property on the shared `WildCard` object, so every module that destructured it sees the new function.
  - Server paths get the patch too: `cfFormulaEngine.ts` and `formulaParserPool.ts` (used by `formulas.ts`) import `formulaOverrides`.
  - The version stays at 1.0.19 through package-lock (package.json still says `^`). `finalReviewC1Wildcard.test.ts` exercises the library path, so a library change would show up there.

**Still open (Important; Critical if edit tokens reach untrusted users).** The matcher is O(key length × pattern length) per key, and nothing caps the total work per request.
- A near-miss pattern (`*` + `a`×290 + `b*`) against 32k-char keys costs about 14 ms per key.
- Measured: 10 criteria × 200 distinct 32k-char keys took **28 s**. The route allows 2,000 criteria per request, which extrapolates to about **5,600 s** in one `POST /api/view/xlsx-cf-ranges`.
- The crafted file is small: 200 cells × 32 KB compresses to a few KB.
- The bake path is mostly bounded by its per-call budget, but one COUNTIF over 1,000 such keys still blocks for about 14 s.
- Suggested fix:
  - Keep a work budget in `countIf`'s key loop (Σ key length × operand length, e.g. 5·10⁷ per request). When it runs out, answer skipped or stale.
  - Or cap wildcard criteria per request (e.g. 32).
  - Optionally use `indexOf` for literal segments.

Two side effects of the library override (informational):
- The library's original regex was unanchored, so `COUNTIF(...,"a*")` also counted "ba". It now does a full match, which is Excel-correct.
- Criteria without the `i` flag stay case-sensitive, as before. Excel is case-insensitive; this was already the case and is unchanged.

### C2 — 반영 (`src/converters/libreOffice.ts:255-310`)

- The attack inputs that took 27 s now take 0–1 ms. The anchor scan also takes 0 ms.
- Differential fuzz over 200,000 random token sequences found **0 differences**. The tokens covered no sectPr, self-closing `<w:sectPr/>`, `<w:sectPrChange>`, unclosed tags, nested `<wp:anchor>`, and `<wp:anchorX>`.
- The `\b` handling in `tagAt` matches the regex, including end of string.
- One divergence remains, and it cannot occur in a valid file. With two `<wp:positionH>` in one anchor, where the first has no `relativeFrom`, the old regex used the second and the new code returns null. The schema allows exactly one positionH with `relativeFrom` required.

### I1 — 반영 (`src/routes/viewer.ts:499-517, 736-742`; `src/docx/import.ts:1686`)

- The unzip is filtered to `word/*.xml` and `*.rels`. I checked that nothing in the docx parse reads media bytes; `mediaPathOf` only builds paths.
- The cache key is `docId:sha1(bytes)`, so a save changes the key and no stale value is served. Old versions age out of the LRU (32 MB, 64 entries).
- Only the Word editor (`editor.ts:341`) uses `docxRender`, and it sends `editor=1`. The hwp, pdf, slides and grid editors and viewer.js do not need it.
- Minor: a docx editor tab opened before the deploy (old bundle, no `editor=1`) gets no `docxRender` until it reloads.

### I2 — 반영, with one gap (`src/collab/room.ts:3057-3116`)

- The size estimate counts `originalPdfBytes`, `pdfEncryptedFile` and `originalXlsxBytes` plus the encoded Y.Doc. A 500 MB PDF room (more than 16 MB) closes at once, as before the grace period.
- `lingeringBytes` stays consistent on all three paths: timer, eviction and wake.
- **New Minor gap.** `lingerBytesOf` runs when the last user leaves. For an unencrypted PDF, `originalPdfBytes` is only set after `await seedPdfModel` (`room.ts:~1866`).
  - If the user closes the tab while a huge PDF is still loading (likely, because huge PDFs load slowly), the estimate is near 0.
  - The room then lingers 45 s, keeps seeding, and holds the whole file.
  - The same applies to a pptx or docx that is still seeding. `pdfParts` is never counted.
  - Fix: do not linger until `room.ready` has settled, or check the size again when seeding ends.

### M1 — 반영 (`room.ts:2522-2524`)

- A no-op `runSerialized` step makes the cleanup wait for any save in progress.
- No deadlock: teardown never awaits from inside a serialized task, and the no-op only joins the end of the chain. A failed save still advances the chain (`previous.then(task, task)`).
- If a new room for the same document queues saves, cleanup is only delayed, and the `!rooms.has` guard still protects that room.
- Unchanged and already present in base: `closeLazyGridStore` (which deletes the row-store file) and `persistChangeLog` still run before that save finishes. The base code had the same order, and the fix did not touch it.

### M2 — 반영 (`room.ts:650-657`)

- Room identity is re-checked after the wait and after counting. An error after close becomes null (409).

### M3 — 반영 (`src/docx/export.ts:329-409, 807-812`)

- Footnote and endnote clone lookup now uses an index built in one indexOf pass. The spread is replaced by a loop. The results match the old regexes for well-formed parts.
- One sibling remains (Minor, not measured, ~70%). `ensureBuiltinLists` (`export.ts:~815`) still uses the tempered regex `(?:(?!</w:abstractNum>)[\s\S])*<w:name w:val=…>`. A numbering.xml with many unclosed `<w:abstractNum>` tags makes saves slow, but only after the user applies a built-in list.

### New defects introduced by the fixes

- None of their own beyond the I2 seeding gap above.
- The C1 residual is not new: it is the original issue reduced from exponential to polynomial.

### Verdict

- C2, I1, M1, M2 and M3 are done.
- C1 is not fully closed: one edit-token request can still block the server for minutes to hours. Before merge, add the per-request work budget (or a cap on wildcard criteria).
- The I2 seeding gap is Minor and can follow later.
