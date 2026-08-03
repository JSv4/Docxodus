# Repo positioning and discoverability

How Docxodus describes itself, and why. This is a working document — when the surface of the library
changes, the copy here is what needs to change with it.

The problem it solves: the repo used to describe itself as *"Office XML Redline Engine Based on
OpenXML SDK (Forked from OpenXMLTools)"*. That is accurate about one of five things the project does,
leads with a fork credit, and contains none of the words people actually search for. Someone looking
for a DOCX viewer, a browser-side Word editor, or a way to let an LLM edit a contract would never
find it, and if they landed on it they would not recognize what they had.

---

## 1. Apply these

These are not in-repo settings — a maintainer has to set them in the GitHub UI (Settings → General,
and the ⚙ next to *About* on the repo home page).

### Repository description

> Swiss-army toolkit for .docx — faithful DOCX→HTML rendering, an in-browser block editor, Word-grade redlines with native tracked changes, and anchor-addressed markdown for LLM agents. .NET, WASM/TypeScript, Python, CLI.

218 characters (GitHub allows 350). It leads with the noun people search for (`.docx`), names all four
capabilities, and closes with the runtimes — which is what a developer scanning search results needs
in order to rule the project in or out.

### Topics

Twenty is GitHub's cap; use all of them, because topic pages are a real discovery path.

```
docx  ooxml  openxml  word  document-comparison  redline  tracked-changes  diff
docx-to-html  html-to-docx  document-editor  markdown  llm  agents  wasm
webassembly  dotnet  csharp  typescript  python
```

Grouped by the audience each is meant to catch:

| Group | Topics | Who finds you |
|---|---|---|
| Format | `docx` `ooxml` `openxml` `word` | Anyone fighting the format at all |
| Compare | `document-comparison` `redline` `tracked-changes` `diff` | Legal tech, contract review, doc management |
| Convert | `docx-to-html` `html-to-docx` | The single highest-volume query in this space |
| Edit | `document-editor` | People building a Word-like surface |
| Agents | `markdown` `llm` `agents` | The newest and fastest-growing segment |
| Runtime | `wasm` `webassembly` `dotnet` `csharp` `typescript` `python` | Language-first searchers |

### Social preview image

Settings → General → Social preview. Without one, every link to the repo — in Slack, on X, in a
newsletter — renders as a grey placeholder. `docs/images/redline.png` cropped to 1280×640 is the
strongest candidate: it is instantly legible as "this diffs contracts" even at thumbnail size.

---

## 2. What is already applied in-repo

| File | Field | Why |
|---|---|---|
| `README.md` | whole file | Rewritten around the four capabilities, each with a real screenshot |
| `npm/package.json` | `description`, `keywords` | npm search ranks heavily on both; the old description mentioned only comparison and conversion |
| `Docxodus/Docxodus.csproj` | `Description`, `PackageTags` | NuGet search ranks on tags; the old description was one sentence about the fork |
| `python/pyproject.toml` | `keywords` | PyPI search |

---

## 3. Keyword strategy

The terms people actually type, and where each one has to appear to be found. A keyword that appears
only in an architecture doc is invisible — GitHub code search indexes it, but repo search and Google
weight the README and the description far more.

| Query cluster | Must appear in |
|---|---|
| `docx to html`, `word to html c#`, `docx viewer javascript` | Description, topics, README H2, npm description |
| `docx diff`, `word document compare`, `redline generator`, `tracked changes api` | Description, topics, README H2, NuGet tags |
| `edit docx in browser`, `docx editor javascript`, `word editor react` | Topics, README H2, npm keywords |
| `llm edit word document`, `docx for agents`, `docx markdown` | Description, topics, README H2, PyPI keywords |
| `openxml sdk alternative`, `openxmlpowertools .net 8` | README body (the fork lineage is a real search path — keep it, just not in the lede) |

Two rules that matter more than the list:

- **Lead with the noun, not the architecture.** "Structure-aware OOXML intermediate representation"
  describes the engine truthfully and matches nothing anyone types. "Compare two Word documents and
  get a redline Word will open" matches a lot.
- **Say what it *is* before what it *forks*.** The OpenXmlPowerTools lineage earns trust and catches
  a genuine migration query, so it stays — at the bottom, where it reads as heritage rather than as
  the headline.

---

## 4. README structure, and why

| Section | Job |
|---|---|
| Tagline + badges | Answer "what is this" in one line and "is it alive" in one glance |
| Positioning paragraph + diagram | Establish that it is one engine with several surfaces, not several tools |
| Four capability sections | Each opens with a **real screenshot**, then explains what the screenshot proves |
| Get started | Four runtimes, four snippets, none longer than six lines |
| Where it runs | The package/install table a developer scrolls to find |
| What else is in the box | The long tail (merge, split, templates, spreadsheets) without burying the headline |
| Documentation | Route the reader into `docs/architecture/` |

**Screenshots carry this README.** A document toolkit is judged on output fidelity, and no prose
substitutes for showing a real redlined contract. Every image is captured from the
[NVCA model financing documents](https://nvca.org/model-legal-documents/) — real forms, not lorem
ipsum — because a synthetic sample invites the reader to assume it only works on synthetic samples.

**Code is kept deliberately thin.** Long inline examples go stale the moment a signature changes and
nobody notices, because a README is not compiled or tested. Snippets show shape only; anything
longer lives in `docs/`, in `npm/examples/`, or in the test suite, where it is exercised.

---

## 5. Keeping it honest

Every claim in the README should be traceable to something checked in:

| Claim | Source |
|---|---|
| Redline change types in the screenshot | The screenshot is the render of a `DocxDiff.Compare` output |
| `accept ≡ right`, `reject ≡ left` | `DocxDiffOps.AcceptRevisions`/`RejectRevisions` + round-trip tests |
| Editor structural-op timings | Measured numbers in `CHANGELOG.md` and `docs/architecture/editor_ui_surface.md` |
| Annotation overlay cost | `docs/architecture/incremental_annotation_overlay.md` |
| Test count | `Docxodus.Tests` xUnit attribute count |

When a number moves, move it here and in the README together. A stale benchmark is worse than no
benchmark.

## 6. Regenerating the screenshots

The images in `docs/images/` are real converter and diff output, captured headlessly:

1. Fetch a real source document (the NVCA model charter and voting agreement are public).
2. Render it with `WmlToHtmlConverter` — flow mode for `render.png`/`footnotes.png`, and with
   `StampAnchors` for the projection figure.
3. For `redline.png`, use the deterministic
   [`tools/screenshots/redline`](../tools/screenshots/redline/README.md) fixture. It applies the
   realistic edit round (fill a blank, tighten wording, insert a definition, move a clause, delete
   a clause), runs `DocxDiff.Compare`, renders with `RenderTrackedChanges: true`, and refuses to
   capture unless the expected cascaded marker sequence is present.
4. The capture script uses the historical 1324×741 CSS viewport at `deviceScaleFactor: 1.4`,
   producing the checked-in 1854×1037 PNG directly.

The editor screenshots come from `npm/examples/editor.html` — see
[`editor_ui_surface.md`](architecture/editor_ui_surface.md).
