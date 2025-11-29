# Move Detection Implementation Plan

## Issue #20: Add move detection to WmlComparer

This document outlines the implementation plan for detecting moved content in the document comparison engine.

---

## Problem Statement

Currently, when content is moved from one location to another in a document, `WmlComparer.GetRevisions()` returns:
1. A `Deleted` revision at the original location
2. An `Inserted` revision at the new location

This makes it impossible for consumers to distinguish between:
- Content that was truly deleted and new content that was added
- Content that was simply relocated within the document

---

## Design Decision: Post-Processing vs. Deep Integration

### Option A: Post-Processing in GetRevisions() ✅ **Recommended**

After collecting deletions and insertions, compare their text content using similarity metrics and link matching pairs as moves.

**Pros:**
- Non-invasive to core comparison algorithm
- Lower risk of regression
- Aligns with the issue's stated approach
- Can be toggled via settings
- Immediate value to API consumers

**Cons:**
- Doesn't generate Word-native `w:moveFrom`/`w:moveTo` markup
- May not catch all edge cases that deep integration would

### Option B: Deep Integration in LCS Algorithm

Modify `ProcessCorrelatedHashes()` to detect moves during the comparison phase and generate `w:moveFrom`/`w:moveTo` markup.

**Pros:**
- Produces Word-native move markup
- More accurate detection
- HTML renderer can display native move formatting

**Cons:**
- High risk of regression
- Complex algorithm changes
- Longer implementation time

**Decision:** Implement Option A for this issue. Option B can be a future enhancement.

---

## Implementation Phases

### Phase 1: Extend Data Structures

#### 1.1 Extend `WmlComparerRevisionType` enum

**File:** `Docxodus/WmlComparer.cs` (line ~3317)

```csharp
public enum WmlComparerRevisionType
{
    Inserted,
    Deleted,
    Moved,  // NEW: Content was relocated (applies to both from/to)
}
```

#### 1.2 Extend `WmlComparerRevision` class

**File:** `Docxodus/WmlComparer.cs` (line ~3323)

```csharp
public class WmlComparerRevision
{
    public WmlComparerRevisionType RevisionType;
    public string Text;
    public string Author;
    public string Date;
    public XElement ContentXElement;
    public XElement RevisionXElement;
    public Uri PartUri;
    public string PartContentType;

    // NEW: Move pair tracking
    /// <summary>
    /// For Moved revisions, this ID links the source and destination.
    /// Both the "from" and "to" revisions share the same MoveGroupId.
    /// </summary>
    public int? MoveGroupId;

    /// <summary>
    /// For Moved revisions: true = this is the source (content moved FROM here),
    /// false = this is the destination (content moved TO here), null = not a move.
    /// </summary>
    public bool? IsMoveSource;
}
```

#### 1.3 Add move detection settings

**File:** `Docxodus/WmlComparer.cs` (line ~52)

```csharp
public class WmlComparerSettings
{
    // ... existing fields ...

    /// <summary>
    /// Whether to detect and mark moved content. Default: true.
    /// When enabled, deletion/insertion pairs with similar text are marked as moves.
    /// </summary>
    public bool DetectMoves = true;

    /// <summary>
    /// Minimum Jaccard similarity (0.0 to 1.0) to consider content as moved.
    /// Default: 0.8 (80% word overlap required).
    /// </summary>
    public double MoveSimilarityThreshold = 0.8;

    /// <summary>
    /// Minimum word count for content to be considered for move detection.
    /// Very short text (< 3 words) is excluded to avoid false positives.
    /// Default: 3.
    /// </summary>
    public int MoveMinimumWordCount = 3;
}
```

---

### Phase 2: Implement Move Detection Logic

#### 2.1 Add similarity calculation method

**File:** `Docxodus/WmlComparer.cs` (new static method)

```csharp
/// <summary>
/// Calculate Jaccard similarity between two strings (word-level).
/// Returns value between 0.0 (no overlap) and 1.0 (identical).
/// </summary>
private static double CalculateJaccardSimilarity(string text1, string text2, WmlComparerSettings settings)
{
    if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
        return 0.0;

    var words1 = TokenizeForComparison(text1, settings);
    var words2 = TokenizeForComparison(text2, settings);

    if (words1.Count == 0 || words2.Count == 0)
        return 0.0;

    var intersection = words1.Intersect(words2).Count();
    var union = words1.Union(words2).Count();

    return union == 0 ? 0.0 : (double)intersection / union;
}

/// <summary>
/// Tokenize text into words for comparison, applying normalization settings.
/// </summary>
private static HashSet<string> TokenizeForComparison(string text, WmlComparerSettings settings)
{
    var separators = settings.WordSeparators ?? new[] { ' ', '-', ')', '(', ';', ',' };
    var words = text.Split(separators, StringSplitOptions.RemoveEmptyEntries);

    if (settings.CaseInsensitive)
        words = words.Select(w => w.ToUpperInvariant()).ToArray();

    return new HashSet<string>(words);
}

/// <summary>
/// Count words in text using the configured word separators.
/// </summary>
private static int CountWords(string text, WmlComparerSettings settings)
{
    if (string.IsNullOrWhiteSpace(text))
        return 0;

    var separators = settings.WordSeparators ?? new[] { ' ', '-', ')', '(', ';', ',' };
    return text.Split(separators, StringSplitOptions.RemoveEmptyEntries).Length;
}
```

#### 2.2 Add move detection post-processing

**File:** `Docxodus/WmlComparer.cs` (new static method)

```csharp
/// <summary>
/// Post-process revisions to detect and mark moved content.
/// Matches deletions with insertions by text similarity.
/// </summary>
private static void DetectMoves(List<WmlComparerRevision> revisions, WmlComparerSettings settings)
{
    if (!settings.DetectMoves)
        return;

    // Separate deletions and insertions
    var deletions = revisions
        .Where(r => r.RevisionType == WmlComparerRevisionType.Deleted)
        .Where(r => !string.IsNullOrWhiteSpace(r.Text))
        .Where(r => CountWords(r.Text, settings) >= settings.MoveMinimumWordCount)
        .ToList();

    var insertions = revisions
        .Where(r => r.RevisionType == WmlComparerRevisionType.Inserted)
        .Where(r => !string.IsNullOrWhiteSpace(r.Text))
        .Where(r => CountWords(r.Text, settings) >= settings.MoveMinimumWordCount)
        .ToList();

    if (deletions.Count == 0 || insertions.Count == 0)
        return;

    int nextMoveGroupId = 1;
    var matchedInsertions = new HashSet<WmlComparerRevision>();

    // For each deletion, find the best matching insertion
    foreach (var deletion in deletions)
    {
        WmlComparerRevision bestMatch = null;
        double bestSimilarity = 0;

        foreach (var insertion in insertions)
        {
            if (matchedInsertions.Contains(insertion))
                continue;

            var similarity = CalculateJaccardSimilarity(deletion.Text, insertion.Text, settings);

            if (similarity >= settings.MoveSimilarityThreshold && similarity > bestSimilarity)
            {
                bestSimilarity = similarity;
                bestMatch = insertion;
            }
        }

        if (bestMatch != null)
        {
            // Mark as move pair
            deletion.RevisionType = WmlComparerRevisionType.Moved;
            deletion.MoveGroupId = nextMoveGroupId;
            deletion.IsMoveSource = true;

            bestMatch.RevisionType = WmlComparerRevisionType.Moved;
            bestMatch.MoveGroupId = nextMoveGroupId;
            bestMatch.IsMoveSource = false;

            matchedInsertions.Add(bestMatch);
            nextMoveGroupId++;
        }
    }
}
```

#### 2.3 Integrate into GetRevisions()

**File:** `Docxodus/WmlComparer.cs` (modify `GetRevisions` method)

Add call to `DetectMoves()` before returning:

```csharp
public static List<WmlComparerRevision> GetRevisions(WmlDocument source, WmlComparerSettings settings)
{
    // ... existing code ...

    var finalRevisionList = mainDocPartRevisionList
        .Concat(footnotesRevisionList)
        .Concat(endnotesRevisionList)
        .ToList();

    // NEW: Post-process to detect moves
    DetectMoves(finalRevisionList, settings);

    return finalRevisionList;
}
```

---

### Phase 3: Update WASM Bridge

#### 3.1 Update RevisionInfo class

**File:** `wasm/DocxodusWasm/DocumentComparer.cs`

Add new properties to the internal `RevisionInfo` class (or create a new response class):

```csharp
class RevisionInfo
{
    public string Author { get; set; }
    public string Date { get; set; }
    public string RevisionType { get; set; }
    public string Text { get; set; }
    public int? MoveGroupId { get; set; }      // NEW
    public bool? IsMoveSource { get; set; }    // NEW
}
```

#### 3.2 Update GetRevisionsJson serialization

**File:** `wasm/DocxodusWasm/DocumentComparer.cs` (modify `GetRevisionsJson`)

```csharp
var response = new RevisionsResponse
{
    Revisions = revisions.Select(r => new RevisionInfo
    {
        Author = r.Author ?? "",
        Date = r.Date ?? "",
        RevisionType = r.RevisionType.ToString(),
        Text = r.Text ?? "",
        MoveGroupId = r.MoveGroupId,        // NEW
        IsMoveSource = r.IsMoveSource       // NEW
    }).ToArray()
};
```

#### 3.3 Update JSON context (if using source generation)

**File:** `wasm/DocxodusWasm/JsonContext.cs` (if exists)

Ensure the new properties are included in serialization.

---

### Phase 4: Update npm/TypeScript Types

#### 4.1 Update RevisionType enum

**File:** `npm/src/types.ts`

```typescript
export enum RevisionType {
  /** Text or content that was added/inserted */
  Inserted = "Inserted",
  /** Text or content that was removed/deleted */
  Deleted = "Deleted",
  /** Text or content that was relocated within the document */
  Moved = "Moved",  // NEW
}
```

#### 4.2 Update Revision interface

**File:** `npm/src/types.ts`

```typescript
export interface Revision {
  author: string;
  date: string;
  revisionType: RevisionType | string;
  text: string;

  // NEW: Move tracking
  /**
   * For moved content, this ID links the source and destination revisions.
   * Both the "from" and "to" revisions share the same moveGroupId.
   * Undefined for non-move revisions.
   */
  moveGroupId?: number;

  /**
   * For moved content: true = source (moved FROM here),
   * false = destination (moved TO here).
   * Undefined for non-move revisions.
   */
  isMoveSource?: boolean;
}
```

#### 4.3 Add type guard for moves

**File:** `npm/src/types.ts`

```typescript
/**
 * Type guard to check if a revision is a move operation.
 * @param revision - The revision to check
 * @returns true if the revision is part of a move
 */
export function isMove(revision: Revision): boolean {
  return revision.revisionType === RevisionType.Moved;
}

/**
 * Type guard to check if a revision is a move source (moved FROM here).
 */
export function isMoveSource(revision: Revision): boolean {
  return isMove(revision) && revision.isMoveSource === true;
}

/**
 * Type guard to check if a revision is a move destination (moved TO here).
 */
export function isMoveDestination(revision: Revision): boolean {
  return isMove(revision) && revision.isMoveSource === false;
}

/**
 * Find the matching pair for a move revision.
 * @param revision - A move revision
 * @param allRevisions - All revisions from the document
 * @returns The matching move revision, or undefined if not found
 */
export function findMovePair(
  revision: Revision,
  allRevisions: Revision[]
): Revision | undefined {
  if (!isMove(revision) || revision.moveGroupId === undefined) {
    return undefined;
  }
  return allRevisions.find(
    (r) =>
      r.moveGroupId === revision.moveGroupId &&
      r.isMoveSource !== revision.isMoveSource
  );
}
```

#### 4.4 Update CompareOptions (optional)

**File:** `npm/src/types.ts`

```typescript
export interface CompareOptions {
  // ... existing options ...

  /** Whether to detect and mark moved content (default: true) */
  detectMoves?: boolean;

  /** Similarity threshold for move detection (0.0-1.0, default: 0.8) */
  moveSimilarityThreshold?: number;

  /** Minimum word count for move detection (default: 3) */
  moveMinimumWordCount?: number;
}
```

#### 4.5 Export new functions

**File:** `npm/src/index.ts`

```typescript
import {
  // ... existing imports ...
  isMove,
  isMoveSource,
  isMoveDestination,
  findMovePair,
} from "./types.js";

export {
  // ... existing exports ...
  isMove,
  isMoveSource,
  isMoveDestination,
  findMovePair,
};
```

---

### Phase 5: Testing

#### 5.1 Create test cases

**File:** `Docxodus.Tests/WmlComparerMoveDetectionTests.cs` (new file)

```csharp
[Fact]
public void DetectMoves_IdenticalText_ShouldMarkAsMoved()
{
    // Create doc1 with paragraph A then B
    // Create doc2 with paragraph B then A (reordered)
    // Compare and verify moves are detected
}

[Fact]
public void DetectMoves_SimilarText_AboveThreshold_ShouldMarkAsMoved()
{
    // Test text with >80% similarity
}

[Fact]
public void DetectMoves_DissimilarText_BelowThreshold_ShouldRemainInsertedDeleted()
{
    // Test text with <80% similarity
}

[Fact]
public void DetectMoves_ShortText_BelowMinimum_ShouldRemainInsertedDeleted()
{
    // Test text with <3 words
}

[Fact]
public void DetectMoves_MultipleMoves_ShouldMatchCorrectly()
{
    // Test multiple move pairs are correctly linked
}

[Fact]
public void DetectMoves_Disabled_ShouldNotMarkMoves()
{
    // Test DetectMoves = false setting
}
```

---

## Summary of Files to Modify

| File | Changes |
|------|---------|
| `Docxodus/WmlComparer.cs` | Extend enum, class, settings; add detection logic |
| `wasm/DocxodusWasm/DocumentComparer.cs` | Update serialization |
| `npm/src/types.ts` | Add `Moved` enum, extend `Revision`, add helpers |
| `npm/src/index.ts` | Export new functions |
| `Docxodus.Tests/WmlComparerMoveDetectionTests.cs` | New test file |
| `CHANGELOG.md` | Document the feature |
| `docs/architecture/wml_comparer_gaps.md` | Mark issue as resolved |

---

## Future Enhancements (Out of Scope for Issue #20)

1. **Generate Word-native move markup** (`w:moveFrom`/`w:moveTo`) in the comparison output for proper rendering in Word
2. **Position-aware matching** - Prefer nearby matches over distant ones
3. **Levenshtein distance** - Alternative similarity metric for partial moves
4. **Move with edits** - Detect content that was both moved and modified

---

## Acceptance Criteria

- [ ] `WmlComparerRevisionType.Moved` enum value exists
- [ ] `WmlComparerRevision.MoveGroupId` and `IsMoveSource` properties exist
- [ ] `WmlComparerSettings` includes move detection options
- [ ] `GetRevisions()` detects moves and marks them appropriately
- [ ] Move pairs share the same `MoveGroupId`
- [ ] WASM bridge serializes new properties
- [ ] TypeScript types include `Moved` and helper functions
- [ ] Unit tests cover move detection scenarios
- [ ] Documentation updated
