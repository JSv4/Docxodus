/**
 * Multi-block selection model for the single-root editor.
 *
 * Maps the native DOM selection (one contenteditable root spans blocks) to the ordered set of
 * editable BODY blocks it covers, with each block's clipped content span. Tables are a selection
 * boundary in v1 — blocks inside a table are excluded so a body selection can't act across cells.
 *
 * Pure addressing/geometry: the DOM read happens here, but the offset helpers are injected from
 * editor.ts (which owns the content-offset space — list markers + bidi marks excluded) to avoid a
 * circular import and keep one definition of "content offset".
 */

const EDITABLE_SELECTOR = '[data-anchor][data-editable="1"]';

export type BlockRole = "only" | "first" | "middle" | "last";

export interface BlockSel {
  el: HTMLElement;
  /** Content span within this block, or null = the whole block is covered. */
  span: { start: number; length: number } | null;
  role: BlockRole;
}

export interface MultiBlockSelection {
  /** The covered body blocks, in document order. */
  blocks: BlockSel[];
  /** True when more than one body block is covered. */
  isMultiBlock: boolean;
}

/** Offset helpers the model needs (defined in editor.ts; passed in to avoid an import cycle). */
export interface SelectionDeps {
  contentOffsetOf: (block: HTMLElement, container: Node, offset: number) => number;
  blockContentText: (block: HTMLElement) => string;
  selectionSpanIn: (block: HTMLElement) => { start: number; length: number } | null;
}

/** The editable body blocks the current selection covers (table cells excluded), or null when the
 *  selection is collapsed / empty / entirely outside the body blocks. */
export function readSelection(
  editRoot: HTMLElement,
  deps: SelectionDeps,
): MultiBlockSelection | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
  const range = sel.getRangeAt(0);
  const all = Array.from(editRoot.querySelectorAll<HTMLElement>(EDITABLE_SELECTOR))
    .filter((b) => !b.closest("table")); // tables are a selection boundary (v1)
  const hit = all.filter((b) => {
    try {
      // comparePoint is robust to a selection boundary that normalized onto a wrapper element
      // rather than a block/text node (Range.intersectsNode misses the end block at a
      // (block, childCount) boundary).
      const startsAfterEnd = range.comparePoint(b, 0) > 0; // block begins after the selection ends
      const endsBeforeStart = range.comparePoint(b, b.childNodes.length) < 0; // block ends before it starts
      return !startsAfterEnd && !endsBeforeStart;
    } catch {
      return false;
    }
  });
  if (hit.length === 0) return null;
  const blocks: BlockSel[] = hit.map((el, i) => ({
    el,
    role: (hit.length === 1
      ? "only"
      : i === 0
      ? "first"
      : i === hit.length - 1
      ? "last"
      : "middle") as BlockRole,
    span: spanFor(el, range, hit.length === 1, deps),
  }));
  return { blocks, isMultiBlock: hit.length > 1 };
}

/** This block's clipped selection span: the first block runs selection-start→end-of-block, middle
 *  blocks are whole (null), the last block runs start-of-block→selection-end, and the only block is
 *  the exact selection. */
function spanFor(
  block: HTMLElement,
  range: Range,
  only: boolean,
  deps: SelectionDeps,
): { start: number; length: number } | null {
  const hasStart = block.contains(range.startContainer);
  const hasEnd = block.contains(range.endContainer);
  if (only || (hasStart && hasEnd)) return deps.selectionSpanIn(block);
  const contentLen = deps.blockContentText(block).length;
  if (hasStart) {
    const start = deps.contentOffsetOf(block, range.startContainer, range.startOffset);
    return { start, length: Math.max(0, contentLen - start) };
  }
  if (hasEnd) {
    const end = deps.contentOffsetOf(block, range.endContainer, range.endOffset);
    return { start: 0, length: end };
  }
  return null; // fully-spanned middle block → whole block
}
