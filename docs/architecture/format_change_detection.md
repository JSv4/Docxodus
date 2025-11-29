# Format Change Detection Architecture

> **Status: IMPLEMENTED** (November 2025)

## Overview

This document describes the architecture for detecting and emitting native Word format change markup (`w:rPrChange`/`w:pPrChange`) in `WmlComparer.Compare()`. When enabled, formatting-only changes (bold, italic, font size, etc.) are tracked as distinct revisions, separate from text insertions and deletions.

## Problem Statement

Currently, when comparing documents:
- Text changes are tracked via `w:ins` and `w:del`
- Move operations are tracked via `w:moveFrom` and `w:moveTo`
- **Formatting-only changes are NOT tracked**

If text remains the same but formatting changes (e.g., making text bold), the current implementation treats this as "equal" content and doesn't emit any revision markup.

## Solution: Following the Move Detection Pattern

The move detection architecture established a successful pattern:
1. **Post-LCS Detection**: Analyze atoms AFTER LCS comparison, BEFORE markup emission
2. **Status Extension**: Add new `CorrelationStatus` values
3. **Native Markup**: Emit Word-native revision elements

Format change detection follows this same pattern.

## Algorithm

### Key Insight: We Already Have Both Document's Data

When atoms are marked as `Equal`, the `FlattenToComparisonUnitAtomList` method already stores:
- `ContentElement` - from the modified document (doc2)
- `ContentElementBefore` - from the original document (doc1)
- `ComparisonUnitAtomBefore` - full atom from original, including `AncestorElements`

The `AncestorElements` array includes the `w:r` (run) element, which contains `w:rPr` (run properties).

### Pipeline

```
Compare(doc1, doc2, settings)
    │
    ├─► Lcs() ─────────────────────────────► Deleted, Inserted, Equal
    │
    ├─► FlattenToComparisonUnitAtomList()
    │       Equal atoms have: ContentElement, ContentElementBefore,
    │                         ComparisonUnitAtomBefore (with AncestorElements)
    │
    │   ╔═══════════════════════════════════════════════════════════╗
    │   ║  DetectMovesInAtomList()                                  ║
    │   ║  Deleted + Inserted → MovedSource/MovedDestination        ║
    │   ╠═══════════════════════════════════════════════════════════╣
    │   ║  DetectFormatChangesInAtomList()    ← NEW                 ║
    │   ║  Equal atoms with different rPr → FormatChanged           ║
    │   ║  - Extract rPr from AncestorElements (w:r element)        ║
    │   ║  - Compare before vs after rPr                            ║
    │   ║  - Store FormatChangeInfo with old/new properties         ║
    │   ╚═══════════════════════════════════════════════════════════╝
    │
    ├─► CoalesceRecurse()
    │       Propagates FormatChanged status and FormatChangeInfo
    │
    ├─► MarkContentAsDeletedOrInsertedTransform()
    │       - MovedSource → w:moveFrom
    │       - MovedDestination → w:moveTo
    │       - FormatChanged → content with w:rPrChange    ← NEW
    │
    └─► WmlDocument with tracked revisions
```

## OpenXML Format Change Markup

### Run Property Change (w:rPrChange)

When text formatting changes (bold, italic, font, etc.):

```xml
<w:r>
  <w:rPr>
    <w:b/>                    <!-- New: bold -->
    <w:i/>                    <!-- New: italic -->
    <w:rPrChange w:id="1" w:author="Author" w:date="2025-01-15T10:30:00Z">
      <w:rPr>
        <!-- Old: no bold, no italic (empty or different properties) -->
      </w:rPr>
    </w:rPrChange>
  </w:rPr>
  <w:t>formatted text</w:t>
</w:r>
```

### Paragraph Property Change (w:pPrChange)

When paragraph formatting changes (alignment, spacing, etc.):

```xml
<w:p>
  <w:pPr>
    <w:jc w:val="center"/>    <!-- New: centered -->
    <w:pPrChange w:id="2" w:author="Author" w:date="2025-01-15T10:30:00Z">
      <w:pPr>
        <w:jc w:val="left"/>  <!-- Old: left aligned -->
      </w:pPr>
    </w:pPrChange>
  </w:pPr>
  <w:r><w:t>text</w:t></w:r>
</w:p>
```

## Implementation Steps

### Step 1: Extend CorrelationStatus Enum

```csharp
public enum CorrelationStatus
{
    Nil,
    Normal,
    Unknown,
    Inserted,
    Deleted,
    Equal,
    Group,
    MovedSource,
    MovedDestination,
    FormatChanged,      // NEW: Text equal, run formatting differs
}
```

### Step 2: Add FormatChangeInfo Class

```csharp
/// <summary>
/// Stores information about formatting changes between original and modified content.
/// </summary>
public class FormatChangeInfo
{
    /// <summary>
    /// Run properties from the original document (before changes).
    /// </summary>
    public XElement OldRunProperties { get; set; }

    /// <summary>
    /// Run properties from the modified document (after changes).
    /// </summary>
    public XElement NewRunProperties { get; set; }

    /// <summary>
    /// List of property names that changed (e.g., "bold", "italic", "fontSize").
    /// </summary>
    public List<string> ChangedProperties { get; set; } = new List<string>();
}
```

### Step 3: Add FormatChangeInfo to ComparisonUnitAtom

```csharp
public class ComparisonUnitAtom : ComparisonUnit
{
    // ... existing properties ...

    public int? MoveGroupId;
    public string MoveName;

    /// <summary>
    /// For format-changed content: stores old and new formatting properties.
    /// </summary>
    public FormatChangeInfo FormatChange;  // NEW
}
```

### Step 4: Implement DetectFormatChangesInAtomList()

```csharp
/// <summary>
/// Analyzes Equal atoms for formatting differences.
/// Converts atoms with different rPr to FormatChanged status.
/// </summary>
private static void DetectFormatChangesInAtomList(
    List<ComparisonUnitAtom> atoms,
    WmlComparerSettings settings)
{
    if (!settings.DetectFormatChanges)
        return;

    foreach (var atom in atoms)
    {
        // Only check Equal atoms that have a "before" reference
        if (atom.CorrelationStatus != CorrelationStatus.Equal)
            continue;
        if (atom.ComparisonUnitAtomBefore == null)
            continue;

        // Extract rPr from both documents
        var oldRPr = GetRunPropertiesFromAtom(atom.ComparisonUnitAtomBefore);
        var newRPr = GetRunPropertiesFromAtom(atom);

        // Compare run properties
        if (!AreRunPropertiesEqual(oldRPr, newRPr, settings))
        {
            atom.CorrelationStatus = CorrelationStatus.FormatChanged;
            atom.FormatChange = new FormatChangeInfo
            {
                OldRunProperties = oldRPr,
                NewRunProperties = newRPr,
                ChangedProperties = GetChangedPropertyNames(oldRPr, newRPr)
            };
        }
    }
}

/// <summary>
/// Extracts the w:rPr element from an atom's ancestor w:r element.
/// </summary>
private static XElement GetRunPropertiesFromAtom(ComparisonUnitAtom atom)
{
    // Find the w:r ancestor element
    var runElement = atom.AncestorElements?.FirstOrDefault(a => a.Name == W.r);
    if (runElement == null)
        return null;

    // Get the rPr child element
    return runElement.Element(W.rPr);
}

/// <summary>
/// Compares two rPr elements for equality, ignoring revision tracking elements.
/// </summary>
private static bool AreRunPropertiesEqual(XElement rPr1, XElement rPr2, WmlComparerSettings settings)
{
    // Normalize: treat null as empty rPr
    var normalized1 = NormalizeRunProperties(rPr1);
    var normalized2 = NormalizeRunProperties(rPr2);

    // Compare the normalized XML
    return XNode.DeepEquals(normalized1, normalized2);
}

/// <summary>
/// Normalizes rPr for comparison by removing revision tracking elements
/// and sorting children consistently.
/// </summary>
private static XElement NormalizeRunProperties(XElement rPr)
{
    if (rPr == null)
        return new XElement(W.rPr);

    // Clone and remove revision tracking elements (rPrChange, etc.)
    var normalized = new XElement(W.rPr,
        rPr.Elements()
           .Where(e => e.Name != W.rPrChange)
           .OrderBy(e => e.Name.LocalName)
           .Select(e => new XElement(e.Name, e.Attributes().OrderBy(a => a.Name.LocalName))));

    return normalized;
}

/// <summary>
/// Returns a list of property names that differ between two rPr elements.
/// </summary>
private static List<string> GetChangedPropertyNames(XElement oldRPr, XElement newRPr)
{
    var changed = new List<string>();

    var oldProps = (oldRPr?.Elements() ?? Enumerable.Empty<XElement>())
        .Where(e => e.Name != W.rPrChange)
        .ToDictionary(e => e.Name, e => e);
    var newProps = (newRPr?.Elements() ?? Enumerable.Empty<XElement>())
        .Where(e => e.Name != W.rPrChange)
        .ToDictionary(e => e.Name, e => e);

    // Find properties that were added, removed, or changed
    var allPropertyNames = oldProps.Keys.Union(newProps.Keys);

    foreach (var propName in allPropertyNames)
    {
        var hasOld = oldProps.TryGetValue(propName, out var oldProp);
        var hasNew = newProps.TryGetValue(propName, out var newProp);

        if (hasOld != hasNew || (hasOld && hasNew && !XNode.DeepEquals(oldProp, newProp)))
        {
            changed.Add(GetFriendlyPropertyName(propName));
        }
    }

    return changed;
}

private static string GetFriendlyPropertyName(XName propName)
{
    return propName.LocalName switch
    {
        "b" => "bold",
        "i" => "italic",
        "u" => "underline",
        "strike" => "strikethrough",
        "sz" => "fontSize",
        "szCs" => "fontSizeComplex",
        "rFonts" => "font",
        "color" => "color",
        "highlight" => "highlight",
        "vertAlign" => "verticalAlign",
        _ => propName.LocalName
    };
}
```

### Step 5: Inject Detection in Pipeline

After `DetectMovesInAtomList`, add:

```csharp
// Detect moves BEFORE markup is written
DetectMovesInAtomList(listOfComparisonUnitAtoms, settings);

// Detect format changes in Equal atoms
DetectFormatChangesInAtomList(listOfComparisonUnitAtoms, settings);
```

### Step 6: Update CoalesceRecurse for FormatChanged Status

Add handling to propagate format change status:

```csharp
if (atom.CorrelationStatus == CorrelationStatus.FormatChanged)
{
    element.Add(new XAttribute("Status", "FormatChanged"));
    if (atom.FormatChange != null)
    {
        // Store old rPr as a child element for later use in markup emission
        if (atom.FormatChange.OldRunProperties != null)
        {
            element.Add(new XAttribute(PtOpenXml.pt + "OldRPr",
                atom.FormatChange.OldRunProperties.ToString(SaveOptions.DisableFormatting)));
        }
    }
}
```

### Step 7: Emit Format Change Markup

In `MarkContentAsDeletedOrInsertedTransform()`:

```csharp
else if (status == "FormatChanged")
{
    // Get the old rPr from the attribute
    var oldRPrString = (string)element.Attribute(PtOpenXml.pt + "OldRPr");
    var oldRPr = oldRPrString != null ? XElement.Parse(oldRPrString) : new XElement(W.rPr);

    // Find or create the new rPr in the run
    var run = element.Name == W.r ? element : element.AncestorsAndSelf(W.r).FirstOrDefault();
    if (run != null)
    {
        var newRPr = run.Element(W.rPr) ?? new XElement(W.rPr);

        // Add rPrChange to track the format change
        newRPr.Add(new XElement(W.rPrChange,
            new XAttribute(W.id, s_MaxId++),
            new XAttribute(W.author, settings.AuthorForRevisions),
            new XAttribute(W.date, settings.DateTimeForRevisions),
            oldRPr));

        // Ensure rPr is in the run
        if (run.Element(W.rPr) == null)
            run.AddFirst(newRPr);
    }

    // Return the element with its content (not wrapped like ins/del)
    return element;
}
```

### Step 8: Add WmlComparerSettings

```csharp
public class WmlComparerSettings
{
    // ... existing settings ...

    /// <summary>
    /// Enable detection of formatting-only changes (bold, italic, font size, etc.).
    /// Default: true
    /// </summary>
    public bool DetectFormatChanges { get; set; } = true;
}
```

### Step 9: Update GetRevisions() for Format Changes

Add a new revision type and detection:

```csharp
public enum WmlComparerRevisionType
{
    Inserted,
    Deleted,
    Moved,
    FormatChanged,  // NEW
}

public class WmlComparerRevision
{
    // ... existing properties ...

    /// <summary>
    /// For FormatChanged revisions: details about what formatting changed.
    /// </summary>
    public FormatChangeDetails FormatChange { get; set; }
}

public class FormatChangeDetails
{
    public Dictionary<string, string> OldProperties { get; set; }
    public Dictionary<string, string> NewProperties { get; set; }
    public List<string> ChangedPropertyNames { get; set; }
}
```

In `GetRevisions()`, detect `w:rPrChange` elements:

```csharp
// Look for rPrChange elements
var rPrChanges = mainXDoc.Descendants(W.rPrChange);
foreach (var rPrChange in rPrChanges)
{
    var revision = new WmlComparerRevision
    {
        RevisionType = WmlComparerRevisionType.FormatChanged,
        Author = (string)rPrChange.Attribute(W.author) ?? "",
        Date = (string)rPrChange.Attribute(W.date) ?? "",
        Text = GetTextFromAncestorRun(rPrChange),
        FormatChange = ExtractFormatChangeDetails(rPrChange)
    };
    revisions.Add(revision);
}
```

## Configuration

### WmlComparerSettings

| Setting | Default | Description |
|---------|---------|-------------|
| `DetectFormatChanges` | `true` | Enable/disable format change detection |

### Disabling Format Detection

When `DetectFormatChanges = false`:
- `DetectFormatChangesInAtomList()` returns immediately
- No `w:rPrChange` markup is generated
- Format-only changes are treated as equal content

## Test Plan

### Unit Tests

1. **Simple bold change** - Text becomes bold
2. **Multiple property changes** - Bold + italic + font size
3. **Property removal** - Bold removed
4. **No format change** - Same formatting, verify no FormatChanged status
5. **Format + text change** - Text changed AND formatted (should be ins/del, not format change)
6. **Paragraph property change** - Alignment change (future enhancement)

### Integration Tests

1. **Word compatibility** - Open result in Word, verify format changes in Track Changes
2. **Round-trip** - Compare → GetRevisions → verify FormatChange populated
3. **Accept revisions** - Verify format changes can be accepted/rejected

## Future Enhancements

### Paragraph Property Changes (Phase 2)

Detect `w:pPr` differences and emit `w:pPrChange`:
- Alignment changes
- Indentation changes
- Spacing changes
- List formatting changes

### Section Property Changes (Phase 3)

Detect `w:sectPr` differences and emit `w:sectPrChange`:
- Page size/orientation
- Margins
- Headers/footers

### Table Property Changes (Phase 4)

Detect `w:tblPr` differences and emit `w:tblPrChange`:
- Table borders
- Cell spacing
- Table width

## Complexity Analysis

| Step | Risk | Complexity | Impact |
|------|------|------------|--------|
| Extend CorrelationStatus | Low | Low | Foundation for format tracking |
| FormatChangeInfo class | Low | Low | Data structure |
| DetectFormatChangesInAtomList | Medium | Medium | rPr comparison logic |
| CoalesceRecurse changes | Low | Low | Status propagation |
| Markup emission | Medium | Medium | Correct w:rPrChange placement |
| GetRevisions() changes | Low | Low | Element detection |

## Summary

This architecture:
1. **Follows the proven move detection pattern** - Same injection point, similar logic flow
2. **Leverages existing infrastructure** - Uses `ContentElementBefore` and `ComparisonUnitAtomBefore`
3. **Produces native Word markup** - Documents show format changes in Word's Track Changes
4. **Is incrementally implementable** - Start with run properties, extend to paragraph/section/table
5. **Maintains backward compatibility** - `DetectFormatChanges = false` preserves existing behavior
