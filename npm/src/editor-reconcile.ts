/**
 * Unit-sequence diff for the editor's incremental structural repaint.
 *
 * After a structural op (insert table/row, footnote, delete block, undo/redo)
 * the editor fetches the session's {@link RenderPlan} (`ListBlocks`) and diffs
 * each container's unit sequence against the DOM's, so unchanged blocks keep
 * their DOM nodes and only changed/created units re-render — replacing the
 * whole-document remount that cost seconds on large documents.
 *
 * Pure functions only: DOM walking, rendering, and fallback policy live in
 * `editor.ts` (`DocxEditor.reconcile`).
 */

/** One top-level render unit: a body block (`p`/`h`/`li`), one whole table
 *  (`tbl`), or one footnote/endnote definition (`fn`/`en`). Mirrors the wire
 *  shape of `DocxSessionBridge.ListBlocks`. `sig` (container units only) is a
 *  content signature: a table or note keeps its unid when content INSIDE it
 *  changes, so the diff tokens must include it or a changed container would be
 *  kept stale. */
export interface RenderUnit {
  id: string;
  kind: string;
  sig?: string;
}

/** The diff token for a unit: unid plus the container content signature. The DOM
 *  side composes the same token from a node's `data-anchor` + `data-render-sig`
 *  stamp, so a container whose inside changed diffs as an in-place substitution. */
export function tokenOf(unit: RenderUnit): string {
  return unidOf(unit.id) + (unit.sig ? "|" + unit.sig : "");
}

/** Per-container render plan — the wire shape of `DocxSessionBridge.ListBlocks`. */
export interface RenderPlan {
  body: RenderUnit[];
  footnotes: RenderUnit[];
  endnotes: RenderUnit[];
}

export interface UnitDiff {
  /** newIndex → oldIndex for units that keep their existing DOM node. */
  keep: Map<number, number>;
  /** Indices into the OLD DOM sequence whose nodes must be removed. */
  removed: number[];
  /** Indices into the NEW plan that need a fresh render. */
  added: number[];
  /**
   * added/removed pairs that are an IN-PLACE change (same number of kept units
   * on each side) — e.g. a text edit re-hashing a block's content-addressed
   * unid. A substituted list item keeps its list position, so it renumbers no
   * siblings and may reconcile where a pure li insert/remove may not.
   */
  substituted: Array<{ oldIndex: number; newIndex: number }>;
}

/** unid ("a1b2…") from a full anchor id ("p:body:a1b2…"). */
export function unidOf(id: string): string {
  return id.substring(id.lastIndexOf(":") + 1);
}

/**
 * LCS diff of the DOM's unit-token sequence against the new plan's. Old entries
 * are tokens (see {@link tokenOf} — bare unids for leaf blocks, `unid|sig` for
 * containers). Equal tokens are interchangeable (content-addressed: equal token
 * ⇒ equal content ⇒ equal rendering, up to position-dependent chrome the caller
 * handles). O(n·m) — sequences are document block counts, a few hundred.
 */
export function diffUnits(oldUnids: string[], newUnits: RenderUnit[]): UnitDiff {
  const n = oldUnids.length;
  const m = newUnits.length;
  const newUnids = newUnits.map(tokenOf);

  // lcs[i][j] = LCS length of oldUnids[i..] vs newUnids[j..]
  const lcs: Int32Array[] = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] =
        oldUnids[i] === newUnids[j]
          ? lcs[i + 1][j + 1] + 1
          : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }

  const keep = new Map<number, number>();
  const removed: number[] = [];
  const added: number[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (oldUnids[i] === newUnids[j]) {
      keep.set(j, i);
      i++;
      j++;
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      removed.push(i++);
    } else {
      added.push(j++);
    }
  }
  while (i < n) removed.push(i++);
  while (j < m) added.push(j++);

  // Substitutions: an added index and a removed index with the same number of
  // kept units before them — an in-place change of one unit.
  const keptBeforeOld = (oi: number): number => {
    let c = 0;
    for (const v of keep.values()) if (v < oi) c++;
    return c;
  };
  const keptBeforeNew = (nj: number): number => {
    let c = 0;
    for (const k of keep.keys()) if (k < nj) c++;
    return c;
  };
  const substituted: Array<{ oldIndex: number; newIndex: number }> = [];
  const usedRemoved = new Set<number>();
  for (const nj of added) {
    const target = keptBeforeNew(nj);
    for (const oi of removed) {
      if (usedRemoved.has(oi)) continue;
      if (keptBeforeOld(oi) === target) {
        substituted.push({ oldIndex: oi, newIndex: nj });
        usedRemoved.add(oi);
        break;
      }
    }
  }

  return { keep, removed, added, substituted };
}

/**
 * Whether this diff must fall back to a full remount:
 *  - a PURE li insert or removal (not an in-place substitution) — sibling list
 *    items renumber without their own XML changing, which a unit diff cannot see;
 *  - total churn above `threshold` (a remount is cheaper and simpler than
 *    rendering half the document block-by-block).
 *
 * `oldKinds[i]` is the kind of the old sequence's i-th unit (the DOM knows it).
 * An li substitution deliberately does NOT force a remount — its list position
 * is unchanged; the CALLER must still verify its list facts (numId/ilvl)
 * survived, and remount when they changed.
 */
export function needsRemount(
  diff: UnitDiff,
  newUnits: RenderUnit[],
  oldKinds: string[],
  threshold = 40,
): boolean {
  if (diff.added.length + diff.removed.length > threshold) return true;
  const subNew = new Set(diff.substituted.map((s) => s.newIndex));
  const subOld = new Set(diff.substituted.map((s) => s.oldIndex));
  for (const j of diff.added) {
    if (!subNew.has(j) && newUnits[j]?.kind === "li") return true;
  }
  for (const i of diff.removed) {
    if (!subOld.has(i) && oldKinds[i] === "li") return true;
  }
  return false;
}
