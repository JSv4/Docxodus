# Native Move Markup Architecture

> **Status: IMPLEMENTED** (November 2025)

## Overview

When `WmlComparerSettings.DetectMoves` is enabled (default: `true`), `WmlComparer.Compare()` produces native Word move tracking markup (`w:moveFrom`/`w:moveTo`) instead of treating relocated content as separate deletions and insertions. This enables Microsoft Word to display moves in its Track Changes panel as relocated content.

## Algorithm

The implementation uses a two-phase approach that preserves the core LCS comparison algorithm:

1. **LCS Phase**: The standard Longest Common Subsequence algorithm identifies `Deleted`, `Inserted`, and `Equal` content
2. **Move Detection Phase**: After LCS but before markup emission, deleted/inserted blocks are analyzed for similarity and converted to moves

### Pipeline

```
Compare(doc1, doc2, settings)
    │
    ├─► CreateComparisonUnitAtomList(doc1)
    ├─► CreateComparisonUnitAtomList(doc2)
    │
    ├─► Lcs() ─────────────────────────────► List<CorrelatedSequence>
    │       (LCS identifies: Deleted, Inserted, Equal)
    │
    ├─► MarkRowsAsDeletedOrInserted()
    │
    ├─► FlattenToComparisonUnitAtomList() ──► List<ComparisonUnitAtom>
    │       Each atom has CorrelationStatus: Deleted | Inserted | Equal
    │
    │   ╔═══════════════════════════════════════════════════════════╗
    │   ║  DetectMovesInAtomList()                                  ║
    │   ║  - Groups consecutive atoms by status into blocks         ║
    │   ║  - Compares deleted vs inserted blocks using Jaccard      ║
    │   ║  - Converts matching pairs to MovedSource/MovedDestination║
    │   ╚═══════════════════════════════════════════════════════════╝
    │
    ├─► AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly()
    │
    ├─► ProduceNewWmlMarkupFromCorrelatedSequence()
    │       └─► CoalesceRecurse()
    │               Builds XElement tree with Status attributes
    │               Propagates MoveGroupId and MoveName for moves
    │
    ├─► MarkContentAsDeletedOrInsertedTransform()
    │       - Status="Deleted" → w:del
    │       - Status="Inserted" → w:ins
    │       - Status="MovedSource" → w:moveFrom with range markers
    │       - Status="MovedDestination" → w:moveTo with range markers
    │
    ├─► FixUpRevisionIds()
    │       Ensures range start/end pairs share the same ID
    │
    └─► WmlDocument with tracked revisions
```

## Move Detection Algorithm

### `DetectMovesInAtomList()`

Located at `WmlComparer.cs:3811`, this method:

1. **Groups atoms into blocks**: Consecutive atoms with the same `CorrelationStatus` (Deleted or Inserted) are grouped into `AtomBlock` objects

2. **Extracts text**: Each block's text is extracted by joining the values of its content elements

3. **Filters by word count**: Blocks with fewer words than `MoveMinimumWordCount` (default: 3) are skipped to avoid false positives

4. **Calculates similarity**: For each deleted block, finds the best matching inserted block using Jaccard word similarity

5. **Marks move pairs**: If similarity meets `MoveSimilarityThreshold` (default: 0.8), both blocks are converted:
   - Deleted atoms → `CorrelationStatus.MovedSource`
   - Inserted atoms → `CorrelationStatus.MovedDestination`
   - Both get the same `MoveGroupId` and `MoveName`

### Jaccard Similarity

The same word-level Jaccard similarity used in `GetRevisions()` post-processing:

```
similarity = |intersection of words| / |union of words|
```

Respects `CaseInsensitive` setting for word comparison.

## OpenXML Move Markup Format

### Move Source (content moved FROM here)

```xml
<w:p>
  <w:moveFromRangeStart w:id="1" w:name="move1" w:author="Author" w:date="2025-01-15T10:30:00Z"/>
  <w:moveFrom w:id="2" w:author="Author" w:date="2025-01-15T10:30:00Z">
    <w:r>
      <w:t>This text was relocated.</w:t>
    </w:r>
  </w:moveFrom>
  <w:moveFromRangeEnd w:id="1"/>
</w:p>
```

### Move Destination (content moved TO here)

```xml
<w:p>
  <w:moveToRangeStart w:id="3" w:name="move1" w:author="Author" w:date="2025-01-15T10:30:00Z"/>
  <w:moveTo w:id="4" w:author="Author" w:date="2025-01-15T10:30:00Z">
    <w:r>
      <w:t>This text was relocated.</w:t>
    </w:r>
  </w:moveTo>
  <w:moveToRangeEnd w:id="3"/>
</w:p>
```

### Key Attributes

| Attribute | Purpose |
|-----------|---------|
| `w:name` | Links source and destination (both use same value, e.g., "move1") |
| `w:id` | Unique document-wide identifier; range start/end pairs share the same ID |
| `w:author` | Author name from `WmlComparerSettings.AuthorForRevisions` |
| `w:date` | Timestamp from `WmlComparerSettings.DateTimeForRevisions` |

## Key Implementation Details

### CorrelationStatus Enum

Extended with two new values:

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
    MovedSource,      // Content moved FROM here
    MovedDestination, // Content moved TO here
}
```

### ComparisonUnitAtom Properties

Added to track move information:

```csharp
public int? MoveGroupId;   // Links source and destination atoms
public string MoveName;     // The w:name attribute value (e.g., "move1")
```

### CoalesceRecurse Changes

When building the XML tree, move status and attributes are propagated:

- `Status="MovedSource"` or `Status="MovedDestination"` attribute
- `pt14:MoveGroupId` and `pt14:MoveName` attributes on content elements

### MarkContentAsDeletedOrInsertedTransform Changes

Handles move status by emitting:

- `w:moveFromRangeStart` / `w:moveFrom` / `w:moveFromRangeEnd` for `MovedSource`
- `w:moveToRangeStart` / `w:moveTo` / `w:moveToRangeEnd` for `MovedDestination`

### FixUpRevisionIds Changes

Updated to maintain ID pairing for range elements:

- Range start elements (`moveFromRangeStart`, `moveToRangeStart`) get assigned a new ID
- Range end elements (`moveFromRangeEnd`, `moveToRangeEnd`) reuse the same ID as their corresponding start element

### Element Lists Updated

- **RecursionElements**: Added `W.moveFrom` and `W.moveTo` (processed like `W.ins`/`W.del`)
- **ElementsToThrowAway**: Added range markers (`moveFromRangeStart`, etc.) - skipped during atom creation
- **InvalidElements**: Removed move elements (now allowed in documents)

## Configuration

### WmlComparerSettings

| Setting | Default | Description |
|---------|---------|-------------|
| `DetectMoves` | `true` | Enable/disable move detection |
| `MoveSimilarityThreshold` | `0.8` | Jaccard similarity threshold (0.0-1.0) |
| `MoveMinimumWordCount` | `3` | Minimum words for move consideration |
| `CaseInsensitive` | `false` | Case-insensitive similarity matching |

### Disabling Move Detection

When `DetectMoves = false`:
- `DetectMovesInAtomList()` returns immediately
- No move markup is generated
- Relocated content appears as separate `w:del` and `w:ins` elements

## GetRevisions() Integration

`GetRevisions()` recognizes native move markup:

1. Detects `w:moveFrom` elements → `WmlComparerRevisionType.Moved` with `IsMoveSource = true`
2. Detects `w:moveTo` elements → `WmlComparerRevisionType.Moved` with `IsMoveSource = false`
3. Extracts `MoveGroupId` from the `pt14:MoveGroupId` attribute to link pairs

## Test Coverage

### Native Move Markup Tests (`WmlComparerMoveDetectionTests.cs`)

| Test | Verifies |
|------|----------|
| `NativeMoveMarkup_ShouldContainMoveFromElement` | `w:moveFrom` elements present |
| `NativeMoveMarkup_ShouldContainMoveToElement` | `w:moveTo` elements present |
| `NativeMoveMarkup_ShouldContainRangeMarkers` | Range start/end elements present and paired |
| `NativeMoveMarkup_ShouldLinkPairsViaNameAttribute` | `w:name` matches between source/dest |
| `NativeMoveMarkup_WhenDisabled_ShouldNotContainMoveElements` | No move elements when disabled |
| `NativeMoveMarkup_ShouldHaveRequiredAttributes` | `w:id`, `w:author`, `w:date` present |
| `NativeMoveMarkup_RangeIdsShouldBeProperlyPaired` | Range start/end IDs match |

## Backward Compatibility

- Existing code calling `Compare()` continues to work unchanged
- `DetectMoves = false` produces identical output as before (only `w:del`/`w:ins`)
- Documents without moves produce identical output
- `GetRevisions()` handles both native move markup and legacy documents
