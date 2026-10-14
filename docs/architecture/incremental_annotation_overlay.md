# Incremental Annotation Overlay

This document describes the incremental annotation overlay system, which enables fast annotation manipulation on pre-converted HTML without re-running the DOCX-to-HTML conversion pipeline.

**Source Files:**
- `Docxodus/ExternalAnnotationProjector.cs` (core projection engine)
- `Docxodus/ExternalAnnotationManager.cs` (annotation set creation, validation, serialization)
- `Docxodus/ExternalAnnotation.cs` (types: `ExternalAnnotationSet`, `ExternalAnnotationProjectionSettings`)
- `wasm/DocxodusWasm/DocumentConverter.cs` (WASM JSExport methods)
- `npm/src/index.ts` (TypeScript wrapper functions)
- `npm/src/types.ts` (TypeScript types)

## Problem Statement

Converting a DOCX file to HTML via `WmlToHtmlConverter` is expensive -- approximately 900ms for a typical document. When annotations change (add, remove, toggle visibility), re-converting the entire document to reflect those changes is wasteful. Most of the conversion cost is in parsing the DOCX package, resolving styles, and building the HTML tree. None of that work changes when a user highlights a new text span or hides a label category.

The incremental annotation overlay eliminates this bottleneck by separating document conversion from annotation rendering. Convert once, then manipulate annotations directly on the HTML string.

## Architecture

### The Overlay Pattern

```
                  DOCX file
                      |
                      v
          +-----------------------+
          | WmlToHtmlConverter    |  ~892ms (one time)
          | (full conversion)     |
          +-----------------------+
                      |
                      v
                 Base HTML  <--- cache this
                      |
        +-------------+-------------+
        |             |             |
        v             v             v
  ProjectAll()   Add()         Remove()
   ~56ms         ~0.3ms        ~18ms
        |             |             |
        v             v             v
              Annotated HTML
```

The base HTML is an immutable reference. Every annotation operation starts from either the base HTML (for full projection) or the current annotated HTML (for incremental add/remove). The annotation projector parses the HTML string as an `XElement` tree, manipulates text nodes to insert wrapper `<span>` elements, and serializes the result back to a string.

### Text-Search-Based Projection

Annotations are projected using **text search**, not byte offsets. The projector:

1. Builds a text map of all text nodes in the HTML `<body>`, recording each node's character offset within the concatenated text.
2. Searches for the annotation's `rawText` in this concatenated HTML text.
3. When found, splits the overlapping text nodes and wraps the annotated portion in a `<span>` with CSS classes and data attributes.
4. After each projection, the text map is rebuilt because the tree has been modified.

This approach is necessary because the offsets in `ExternalAnnotationSet` refer to the source document text (from `OpenContractExporter`), which may differ from the HTML text due to whitespace normalization, element boundaries, and content that the HTML converter omits or transforms.

### Offset-Drift Fix

When an annotation is projected, the wrapper `<span>` may include a label child (e.g., `<span class="ext-annot-label">Clause</span>`). The label text would pollute the offset calculation for subsequent annotations if it were included in the text map. The `GetTextNodes` method handles this by skipping elements that have a `data-annotation-id` attribute:

```csharp
// Skip already-projected annotation wrappers so their label
// text doesn't shift offsets during subsequent projections
if (child.Attribute("data-annotation-id") != null)
    continue;
```

This means that after projecting annotation A, the text map for annotation B still reflects the original document text positions.

## API Surface

### .NET (`ExternalAnnotationProjector`)

| Method | Description |
|--------|-------------|
| `ProjectAnnotationsOntoHtml(html, annotationSet, settings?)` | Project all annotations from a set onto an HTML string. Returns annotated HTML. |
| `AddAnnotationToHtml(html, annotation, label?, settings?)` | Add a single annotation to existing HTML. Does not require the full annotation set. |
| `RemoveAnnotationFromHtml(html, annotationId, cssClassPrefix?)` | Remove a single annotation by ID. Unwraps `<span>` elements back to plain text. |
| `GenerateVisibilityCss(hiddenLabelIds, cssClassPrefix?)` | Generate CSS rules that hide annotations with specific label IDs (transparency + `display: none` on labels). |
| `GenerateAnnotationCssString(labels, settings?)` | Generate the full annotation stylesheet for a set of label definitions. |
| `ProjectAnnotations(htmlElement, annotationSet, settings)` | Lower-level: operates on `XElement` instead of string. Used internally. |
| `ConvertWithAnnotations(doc, annotationSet, htmlSettings?, projectionSettings?)` | Convenience: full DOCX conversion + annotation projection in one call. |

### WASM (`DocumentConverter` JSExport methods)

All WASM methods accept and return JSON strings. Responses are wrapped in `HtmlConversionResponse` or `CssResponse` objects.

| Method | Parameters | Returns |
|--------|------------|---------|
| `ProjectAnnotationsOntoHtml` | `html`, `annotationSetJson`, `extAnnotCssClassPrefix`, `extAnnotLabelMode` (int) | `{ html: string }` |
| `AddAnnotationToHtml` | `html`, `annotationJson`, `labelJson`, `extAnnotCssClassPrefix`, `extAnnotLabelMode` (int) | `{ html: string }` |
| `RemoveAnnotationFromHtml` | `html`, `annotationId`, `extAnnotCssClassPrefix` | `{ html: string }` |
| `GenerateAnnotationVisibilityCss` | `hiddenLabelIdsJson` (string[]), `extAnnotCssClassPrefix` | `{ css: string }` |
| `GenerateAnnotationCss` | `labelsJson` (Record<string, AnnotationLabel>), `extAnnotCssClassPrefix`, `extAnnotLabelMode` (int) | `{ css: string }` |

The `extAnnotLabelMode` parameter maps to the `AnnotationLabelMode` enum: `Above = 0`, `Inline = 1`, `Tooltip = 2`, `None = 3`.

### npm/TypeScript

| Function | Signature |
|----------|-----------|
| `projectAnnotationsOntoHtml` | `(html: string, annotationSet: ExternalAnnotationSet, projectionOptions?: ExternalAnnotationProjectionSettings) => Promise<string>` |
| `addAnnotationToHtml` | `(html: string, annotation: OpenContractsAnnotation, label?: AnnotationLabel, projectionOptions?: ExternalAnnotationProjectionSettings) => Promise<string>` |
| `removeAnnotationFromHtml` | `(html: string, annotationId: string, cssClassPrefix?: string) => Promise<string>` |
| `generateAnnotationVisibilityCss` | `(hiddenLabelIds: string[], cssClassPrefix?: string) => Promise<string>` |
| `generateAnnotationCss` | `(labels: Record<string, AnnotationLabel>, projectionOptions?: ExternalAnnotationProjectionSettings) => Promise<string>` |

## How It Works

### 1. Text Map Construction

`BuildTextMap` traverses the HTML body and collects every `XText` node along with its character offset in the concatenated body text. The result is a list of `TextMapEntry` objects:

```
TextMapEntry { TextNode: "This is a ", StartOffset: 0,  EndOffset: 10 }
TextMapEntry { TextNode: "contract",   StartOffset: 10, EndOffset: 18 }
TextMapEntry { TextNode: " between",   StartOffset: 18, EndOffset: 26 }
```

### 2. Text Search

`FindTextInHtml` searches the concatenated text for the annotation's `rawText`. It tracks used offsets to handle duplicate text -- if the same phrase appears multiple times, each annotation claims a distinct occurrence.

### 3. Node Splitting and Wrapping

`WrapTextNode` splits a text node into up to three parts: before, annotated, after. The annotated part is wrapped in a `<span>`:

```html
<!-- Before -->
<span style="...">This is a contract between parties.</span>

<!-- After projecting "contract between" -->
<span style="...">This is a </span>
<span class="ext-annot-highlight ext-annot-single"
      data-annotation-id="ann-001"
      data-label-id="TERM"
      data-label="Term"
      style="--annot-color: #2196F3;">
  <span class="ext-annot-label">Term</span>
  contract between
</span>
<span style="..."> parties.</span>
```

Multi-node annotations (text spanning across elements) produce multiple wrapper spans with position classes: `ext-annot-start`, `ext-annot-continuation`, `ext-annot-end`. The label is rendered only on the first segment.

### 4. CSS Generation

`BuildAnnotationCssString` generates base styles for all annotations plus per-label color classes. Colors are applied through CSS custom properties (`--annot-color`), allowing label-specific styling without unique class names per annotation instance.

`GenerateVisibilityCss` produces override rules that hide annotations by label ID. This allows toggling label visibility purely through CSS, without modifying the HTML.

### 5. Removal

`RemoveAnnotationFromHtml` finds all `<span>` elements with `data-annotation-id` matching the target ID. For each:
1. Remove child label `<span>` elements.
2. Move the remaining child nodes before the wrapper.
3. Remove the now-empty wrapper.

## Performance

Benchmarks from CI (representative document):

| Operation | Time | Speedup vs Full Conversion |
|-----------|------|---------------------------|
| Full DOCX re-conversion | ~892ms | 1x (baseline) |
| Incremental projection (all annotations) | ~56ms | 15.9x faster |
| Single annotation add | ~0.3ms | 2,972x faster |
| Single annotation remove | ~18ms | 49x faster |

The remove operation is slower than add because it requires parsing the full HTML string into an `XElement` tree, searching for matching spans, unwrapping them, and re-serializing. The add operation also parses/serializes but operates on a smaller search scope.

## Typical Usage Pattern

### TypeScript (npm)

```typescript
import {
  convertDocxToHtml,
  createExternalAnnotationSet,
  projectAnnotationsOntoHtml,
  addAnnotationToHtml,
  removeAnnotationFromHtml,
  generateAnnotationVisibilityCss,
  createAnnotation,
} from "docxodus";

// Step 1: Convert once and cache
const baseHtml = await convertDocxToHtml(docxBytes);
const annotationSet = await createExternalAnnotationSet(docxBytes, "doc-123");

// Step 2: Define labels
annotationSet.textLabels["CLAUSE"] = {
  id: "CLAUSE",
  text: "Clause",
  color: "#FF5722",
};
annotationSet.textLabels["TERM"] = {
  id: "TERM",
  text: "Term",
  color: "#2196F3",
};

// Step 3: Create annotations and project all at once
const ann1 = createAnnotation("ann-1", "CLAUSE", annotationSet.content, 100, 250);
const ann2 = createAnnotation("ann-2", "TERM", annotationSet.content, 300, 320);
annotationSet.labelledText.push(ann1, ann2);

let html = await projectAnnotationsOntoHtml(baseHtml, annotationSet);

// Step 4: Incrementally add one more
const ann3 = createAnnotation("ann-3", "TERM", annotationSet.content, 500, 530);
const termLabel = annotationSet.textLabels["TERM"];
html = await addAnnotationToHtml(html, ann3, termLabel);

// Step 5: Remove one
html = await removeAnnotationFromHtml(html, "ann-1");

// Step 6: Toggle visibility by label (CSS only, no HTML change)
const hideCss = await generateAnnotationVisibilityCss(["TERM"]);
// Apply hideCss to a <style> element in the DOM
```

### C# (.NET)

```csharp
// Convert once
var htmlSettings = new WmlToHtmlConverterSettings();
var baseHtml = WmlToHtmlConverter.ConvertToHtml(doc, htmlSettings).ToString();

// Create annotation set
var annotationSet = ExternalAnnotationManager.CreateAnnotationSet(doc, "doc-123");
annotationSet.TextLabels["CLAUSE"] = new AnnotationLabel
{
    Id = "CLAUSE", Text = "Clause", Color = "#FF5722"
};

// Create annotations
var ann = ExternalAnnotationManager.CreateAnnotation(
    "ann-1", "CLAUSE", annotationSet.Content, 100, 250);
annotationSet.LabelledText.Add(ann);

// Project all
string annotatedHtml = ExternalAnnotationProjector.ProjectAnnotationsOntoHtml(
    baseHtml, annotationSet);

// Incremental add
var ann2 = ExternalAnnotationManager.CreateAnnotation(
    "ann-2", "CLAUSE", annotationSet.Content, 300, 350);
annotatedHtml = ExternalAnnotationProjector.AddAnnotationToHtml(
    annotatedHtml, ann2, annotationSet.TextLabels["CLAUSE"]);

// Incremental remove
annotatedHtml = ExternalAnnotationProjector.RemoveAnnotationFromHtml(
    annotatedHtml, "ann-1");

// Visibility toggle
string css = ExternalAnnotationProjector.GenerateVisibilityCss(
    new[] { "CLAUSE" });
```

## Limitations and Caveats

1. **Text-search based, not byte-offset based.** The projector finds annotations by searching for their `rawText` in the HTML text content. If the same text appears multiple times, the projector uses the first unused occurrence. This can cause misalignment if annotations target a later occurrence of identical text and an earlier occurrence is not also annotated.

2. **Annotations must match text in the HTML.** If the HTML converter transforms, omits, or normalizes text differently from the source document, the text search may fail and the annotation will be silently skipped. Examples: collapsed whitespace, special characters converted to entities, content in headers/footers/footnotes that may or may not be present depending on converter settings.

3. **XHTML namespace considerations.** The `ExternalAnnotationProjector` creates wrapper `<span>` elements without a namespace (plain HTML). The `WmlToHtmlConverter` produces XHTML with the `http://www.w3.org/1999/xhtml` namespace. `XElement.Parse` handles this transparently for round-tripping, but consumers that manipulate the tree directly should be aware of the namespace difference.

4. **Full re-parse on every operation.** Each call to `AddAnnotationToHtml`, `RemoveAnnotationFromHtml`, or `ProjectAnnotationsOntoHtml` parses the entire HTML string into an `XElement` tree and serializes it back. For very large documents, this parse/serialize overhead dominates. The performance numbers above include this cost.

5. **No overlapping annotation merge.** When two annotations cover the same text range, they produce nested `<span>` elements. The CSS handles this visually, but deeply nested annotations (many overlapping spans on the same text) can produce verbose HTML.

6. **Annotation labels pollute text content.** When `LabelMode` is `Above` or `Inline`, the label text (e.g., "Clause") is rendered as a child `<span>` inside the annotation wrapper. The `GetTextNodes` skip logic prevents this from affecting offset calculations, but consumers that extract text from the HTML should be aware that label text is present in the DOM.

7. **Remove does not coalesce adjacent text nodes.** After removing an annotation, the text that was split into before/annotated/after nodes is not re-merged. This is functionally correct (the text renders identically) but means the DOM has more text nodes than the original.
