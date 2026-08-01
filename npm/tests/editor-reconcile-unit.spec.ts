import { test, expect } from "@playwright/test";
import { diffUnits, needsRemount, unidOf } from "../src/editor-reconcile";
import type { RenderUnit } from "../src/editor-reconcile";

const U = (ids: string) => ids.split(" ").filter(Boolean);
const P = (ids: string): RenderUnit[] => U(ids).map((u) => ({ id: `p:body:${u}`, kind: "p" }));

test.describe("editor-reconcile unit diff", () => {
  test("insertion keeps all existing nodes", () => {
    const d = diffUnits(U("a b c"), P("a b x c"));
    expect([...d.keep.entries()].sort((x, y) => x[0] - y[0])).toEqual([
      [0, 0],
      [1, 1],
      [3, 2],
    ]);
    expect(d.added).toEqual([2]);
    expect(d.removed).toEqual([]);
  });

  test("change = remove+add at same position, classified as substitution", () => {
    const d = diffUnits(U("a b c"), P("a B c"));
    expect(d.added).toEqual([1]);
    expect(d.removed).toEqual([1]);
    expect(d.keep.size).toBe(2);
    expect(d.substituted).toEqual([{ oldIndex: 1, newIndex: 1 }]);
  });

  test("pure removal", () => {
    const d = diffUnits(U("a b c"), P("a c"));
    expect(d.added).toEqual([]);
    expect(d.removed).toEqual([1]);
    expect(d.keep.size).toBe(2);
  });

  test("duplicate unids treated positionally", () => {
    const d = diffUnits(U("a x x b"), P("a x b"));
    expect(d.removed.length).toBe(1);
    expect(d.added).toEqual([]);
    expect(d.keep.size).toBe(3);
  });

  test("empty old sequence — everything added", () => {
    const d = diffUnits([], P("a b"));
    expect(d.added).toEqual([0, 1]);
    expect(d.removed).toEqual([]);
  });

  test("needsRemount: li creation forces remount, p creation does not", () => {
    const oldUnids = U("a");
    const oldKinds = ["p"];
    const withLi: RenderUnit[] = [
      { id: "p:body:a", kind: "p" },
      { id: "li:body:n", kind: "li" },
    ];
    expect(needsRemount(diffUnits(oldUnids, withLi), withLi, oldKinds)).toBe(true);
    const withP = P("a n");
    expect(needsRemount(diffUnits(oldUnids, withP), withP, oldKinds)).toBe(false);
  });

  test("needsRemount: li removal forces remount", () => {
    const oldUnids = U("a n");
    const oldKinds = ["p", "li"];
    const next = P("a");
    expect(needsRemount(diffUnits(oldUnids, next), next, oldKinds)).toBe(true);
  });

  test("needsRemount: p removal does not force remount", () => {
    const oldUnids = U("a n");
    const oldKinds = ["p", "p"];
    const next = P("a");
    expect(needsRemount(diffUnits(oldUnids, next), next, oldKinds)).toBe(false);
  });

  test("needsRemount: churn above threshold", () => {
    const many = P(Array.from({ length: 50 }, (_, i) => `n${i}`).join(" "));
    expect(needsRemount(diffUnits([], many), many, [])).toBe(true);
    const few = P(Array.from({ length: 5 }, (_, i) => `n${i}`).join(" "));
    expect(needsRemount(diffUnits([], few), few, [])).toBe(false);
  });

  test("li SUBSTITUTION (text edit re-hashes the unid) does not force remount", () => {
    // A text edit on a list item changes its content-hashed unid: the diff sees
    // remove+add at one position. Its list position is unchanged, so no sibling
    // renumbers — the reconciler swaps it in place (facts check is the caller's).
    const oldUnids = U("a n b");
    const oldKinds = ["p", "li", "p"];
    const next: RenderUnit[] = [
      { id: "p:body:a", kind: "p" },
      { id: "li:body:N2", kind: "li" },
      { id: "p:body:b", kind: "p" },
    ];
    const d = diffUnits(oldUnids, next);
    expect(d.substituted).toEqual([{ oldIndex: 1, newIndex: 1 }]);
    expect(needsRemount(d, next, oldKinds)).toBe(false);
  });

  test("li surviving unchanged does not force remount", () => {
    const oldUnids = U("a n b");
    const oldKinds = ["p", "li", "p"];
    const next: RenderUnit[] = [
      { id: "p:body:a", kind: "p" },
      { id: "li:body:n", kind: "li" },
      { id: "p:body:B", kind: "p" },
    ];
    expect(needsRemount(diffUnits(oldUnids, next), next, oldKinds)).toBe(false);
  });

  test("unidOf", () => {
    expect(unidOf("p:body:abc123")).toBe("abc123");
    expect(unidOf("tbl:body:ff00")).toBe("ff00");
  });
});
