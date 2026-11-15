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
docx-to-html  document-editor  markdown  llm  ai-agents  mcp  mcp-server
wasm  dotnet  csharp  typescript  python
```

`mcp` and `mcp-server` are how people and agent-tool directories browse for servers, and
`ai-agents` is the busier spelling of `agents`. They take the slots of `html-to-docx` (a minor
capability), `webassembly` (`wasm` already covers it) and `agents`.

Grouped by the audience each is meant to catch:

| Group | Topics | Who finds you |
|---|---|---|
| Format | `docx` `ooxml` `openxml` `word` | Anyone fighting the format at all |
| Compare | `document-comparison` `redline` `tracked-changes` `diff` | Legal tech, contract review, doc management |
| Convert | `docx-to-html` | The single highest-volume query in this space |
| Edit | `document-editor` | People building a Word-like surface |
| Agents | `markdown` `llm` `ai-agents` `mcp` `mcp-server` | The newest and fastest-growing segment |
| Runtime | `wasm` `dotnet` `csharp` `typescript` `python` | Language-first searchers |

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
| `python/pyproject.toml` | `description`, `keywords`, `classifiers`, `urls` | PyPI search; the project page's sidebar links |
| `npm-export/package.json` | `description`, `keywords` | "docx to pdf" is a query of its own, and the old description named neither format's source |
| `tools/*/*.csproj` | `Description`, `PackageTags` | Each CLI tool is its own NuGet search result |
| `docs/llms.txt` | whole file | The index an agent reads — see section 4 |
| `docs/demo/index.html` | title, description, canonical URL, JSON-LD | The one page of ours a search engine indexes as a site rather than as a repository |
| `docs/sitemap.xml` | whole file | Lists the demo pages and `llms.txt` |

`docs/demo/tools/discoverability.test.mjs` (run by `npm run test:demo-logic`) keeps these from
going stale: every project link in `llms.txt` and the sitemap must resolve to a checked-in file,
every API `llms.txt` names must exist in the source of the surface it is named for, the MCP tool
count must match the catalog, and no package may advertise a format the library dropped — the
NuGet tags listed `xlsx` and `pptx` for several releases after that support was removed.

Two limits of a project site on `github.io` worth knowing: a `robots.txt` is only honoured at the
host root, which belongs to the account, not this repository, so there is none here and the
sitemap has to be submitted by hand (Google Search Console → Sitemaps) to be read; and `llms.txt`
is served at `/Docxodus/llms.txt`, not the host root, so it is found through the links to it — the
README, the landing page's `<link rel="alternate">`, the sitemap — not by convention.

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
| `docx javascript`, `docx diff nodejs`, `compare word documents typescript` | README lede ("It's JavaScript"), npm description and keywords, `llms.txt` |
| `llm edit word document`, `docx for agents`, `docx markdown`, `docx mcp server` | Description, topics, README H2 ("For AI agents"), PyPI keywords, `llms.txt` |
| `blackline`, `compare contracts`, `docx to pdf node` | README body, NuGet/npm tags, npm-export description |
| `openxml sdk alternative`, `openxmlpowertools .net 8` | README body (the fork lineage is a real search path — keep it, just not in the lede) |

Two rules that matter more than the list:

- **Say JavaScript, not WebAssembly.** A JavaScript developer who reads ".NET compiled to WASM"
  assumes a toolchain to set up. The npm package is ordinary ES modules with types; WebAssembly is
  how it works, not what it asks of the reader, so it goes second wherever both appear.
- **Lead with the noun, not the architecture.** "Structure-aware OOXML intermediate representation"
  describes the engine truthfully and matches nothing anyone types. "Compare two Word documents and
  get a redline Word will open" matches a lot.
- **Say what it *is* before what it *forks*.** The OpenXmlPowerTools lineage earns trust and catches
  a genuine migration query, so it stays — at the bottom, where it reads as heritage rather than as
  the headline.

---

## 4. Being found by agents

A person skims a README; an agent deciding which library to use reads whatever a search hands it
and then looks for the call to make. Three things serve that reader:

- **`docs/llms.txt`** ([llmstxt.org](https://llmstxt.org) format) says what the project is in one
  paragraph, then maps each task — compare, render, convert, edit, accept and reject changes — to
  the exact call on every surface. Its links go to raw markdown, which a model reads without
  scraping GitHub's HTML. Write it in the words of the task ("compare two .docx files into a
  tracked-changes redline"), not of the architecture.
- **The README's "For AI agents" section** routes an agent to the right surface and carries a
  copy-paste MCP configuration.
- **Tool descriptions in `tools/mcp-server/ToolCatalog.cs`** are what an agent sees once
  connected; they are the product copy for that audience.

**The open gap: the MCP server is not published.** `Docxodus.McpServer` packs as a .NET tool but
`publish.yml` does not push it, so an agent cannot install it by name and no MCP directory can
list it. Publishing it to NuGet — whose search has a dedicated MCP-server package type, read from a
`.mcp/server.json` in the package — and then registering it with the
[MCP registry](https://github.com/modelcontextprotocol/registry) would do more for agent discovery
than everything else in this section. It is a release-process decision (a fifth NuGet package, a
~160 MB self-contained tool), so it is recorded here rather than done.

---

## 5. README structure, and why

| Section | Job |
|---|---|
| Tagline + badges | Answer "what is this" in one line and "is it alive" in one glance |
| Positioning paragraph + diagram | Establish that it is one engine with several surfaces, not several tools |
| Four capability sections | Each opens with a **real screenshot**, then explains what the screenshot proves |
| For AI agents | Route an agent to a surface; MCP configuration; link to `llms.txt` |
| Get started | Four runtimes, four snippets, none longer than six lines |
| Where it runs | The package/install table a developer scrolls to find |
| What else is in the box | The long tail (merge, split, templates, regex replace) without burying the headline |
| Documentation | Route the reader into `docs/architecture/` |

**Screenshots carry this README.** A document toolkit is judged on output fidelity, and no prose
substitutes for showing a real redlined contract. Every image is captured from the
[NVCA model financing documents](https://nvca.org/model-legal-documents/) — real forms, not lorem
ipsum — because a synthetic sample invites the reader to assume it only works on synthetic samples.

**Code is kept deliberately thin.** Long inline examples go stale the moment a signature changes and
nobody notices, because a README is not compiled or tested. Snippets show shape only; anything
longer lives in `docs/`, in `npm/examples/`, or in the test suite, where it is exercised.

---

## 6. Keeping it honest

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

## 7. Regenerating the screenshots

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
