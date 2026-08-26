# Findings from the 2026-08-26 run (NVCA Model COI, Oct 2025)

Measured at commit c8e13d2 on .NET 10.0.400, single container, cold process per phase.
Timings are indicative, not benchmarks-grade.

## Headline results

| Measurement | Result |
|---|---|
| HTML (footnotes + headers rendered) | 0.8–2.1 s |
| Markdown projection | 364–459 ms, 214,469 chars, 462 anchors, all 94 footnotes projected inline |
| No-edit session round-trip | 401–535 ms, text-exact, part inventory identical |
| Tracked session, 8-edit counsel script | all edits succeed; 11 native revisions, correct author; single accept/reject verified |
| `DocxDiff.Compare` | 1.5–2.4 s, 18 word-level revisions with both-side anchors |
| Accept-all / reject-all round-trip | exact in both directions |
| Schema findings added by any output | 0 (source baseline 80 → outputs 80) |
| Edit-script / semantic-changes JSON | 46.6 KB / ~50 KB, 24 typed changes with SHA-256 content+format digests |
| Redline → HTML (tracked markup) | 0.7–0.8 s, semantic `<ins>`/`<del>` |
| `@docxodus/export` PDF | 51 pages in 17.3 s cold (Chromium launch included), render report `complete`; redline with `reviewProfile: "markup"` renders as Word-style strikeout/underline markup |
| `docxodus-mcp` end-to-end (open → search → tracked edit → revision list → sessionless compare) | first-try success; compare 3.8 s with per-author revision summary |

## Issues discovered (Docxodus)

1. **`WmlComparer` (legacy engine) fails its round-trip invariants on this document.**
   Accept-all and reject-all both diverge from the expected text: the engine silently
   drops an empty paragraph inside a footnote (adjacent to the *Kumar v. Racing Corp.*
   footnote), and it rewrites the package wholesale (validator findings drop 80 → 29,
   i.e. it normalizes markup it should preserve). It is also ~3× slower than `DocxDiff`
   (6.3–7.5 s vs 1.5–2.4 s). The benchmark keeps the legacy engine as a tracked
   known-gap: the two `[check] legacy ... FAIL` lines are expected until this is fixed
   or the engine's documentation explicitly scopes it away from footnote-heavy
   documents. `DocxDiff` on the identical inputs passes both invariants exactly.

2. **`DocxSession.DeleteBlock` leaves orphaned footnote definitions.** Deleting a
   paragraph whose text carries a footnote reference removes the reference but leaves
   the footnote body in `word/footnotes.xml`. Word renders nothing (the note is
   unreferenced), but the text still ships inside the file — for legal workflows this is
   a confidentiality-adjacent leak: "deleted" drafting commentary survives in the
   package. Options: prune unreferenced notes on delete, prune on `Save()`, or expose a
   `CompactFootnotes`-style op; whichever is chosen should be revision-aware (a tracked
   delete must keep the note until the revision is accepted).

3. **Formatting-only edits that cross a field envelope surface as delete+insert pairs
   in the native redline.** Italicizing a span that intersects a cross-reference field
   produces a full del+ins of the field region rather than a formatting revision —
   correct OOXML, but noisy for a human reviewer. The semantic changeset already
   classifies it precisely (`run_formatting` modify on the exact token span plus a
   `field` envelope change), so this is a markup-shaping improvement, not a diff bug.

4. **PDF export's sandbox posture will surprise container deployments.**
   `@docxodus/export` (correctly) refuses to launch Chromium without its OS sandbox: it
   fails when run as root and needs unprivileged user namespaces enabled. The error and
   remediation text are good; the README warning deserves to be louder, and a
   preflight `checkEnvironment()` helper would turn the first failed conversion into a
   configuration message.

5. **Small DX nits.** `DocxDiffRevision` has no useful `ToString()` (logging a revision
   prints the type name); the projection-side tracked-changes knob
   (`ProjectionSettings.TrackedChanges`) is separate from the mutation-recording knob
   (`SetTrackedChanges`) and easy to conflate — an agent that records tracked edits and
   then projects sees clean text unless it also sets the projection mode.
