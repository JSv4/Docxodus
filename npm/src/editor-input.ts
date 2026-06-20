/**
 * beforeinput classifier for the single-root editor.
 *
 * Given a beforeinput inputType and whether the current selection spans multiple body blocks, decide
 * what the editor should do. Pure (no DOM access) so the routing rules are unit-reasonable and the
 * editor stays the only place that touches the session.
 *
 * Single-block input is left NATIVE (the browser edits the text node; the change commits on
 * blur/selection-leave). Only structural / cross-block operations are intercepted and routed to a
 * compound handler. Native rich-text formatting input (Ctrl+B etc. that the browser would turn into
 * <b>/<i>) is always rerouted to the editor's lossless format() so no stray markup is injected.
 */

export type InputAction =
  | { kind: "native" } //               let the browser do it (single-block text / IME)
  | { kind: "deleteSelection" } //      collapse a multi-block selection (Backspace/Delete/cut)
  | { kind: "typeOver"; text: string } // replace a multi-block selection with typed text
  | { kind: "splitAtSelection" } //     Enter over a multi-block selection
  | { kind: "paste" } //                paste over a multi-block selection (plain text, v1)
  | { kind: "format"; key: string } //  native formatBold/Italic/… → our format()
  | { kind: "block" }; //               disabled in v1 (drag-and-drop of content)

const FORMAT_MAP: Record<string, string> = {
  formatBold: "bold",
  formatItalic: "italic",
  formatUnderline: "underline",
  formatStrikeThrough: "strike",
  formatSuperscript: "superscript",
  formatSubscript: "subscript",
};

export function classifyBeforeInput(
  inputType: string,
  data: string | null,
  isMultiBlock: boolean,
): InputAction {
  // Native rich-text formatting is always rerouted (single- or multi-block) so the browser never
  // injects its own <b>/<i> markup — the editor applies formatting losslessly via the session.
  if (FORMAT_MAP[inputType]) return { kind: "format", key: FORMAT_MAP[inputType] };
  // Cross-block drag-and-drop of content is out of scope for v1; disable it so it can't corrupt the
  // model. (Single-block drag-drop is rare and also disabled — acceptable for v1.)
  if (inputType === "deleteByDrag" || inputType === "insertFromDrop") return { kind: "block" };

  // Everything else only needs special handling when the selection spans more than one body block.
  if (!isMultiBlock) return { kind: "native" };

  switch (inputType) {
    case "insertText":
    case "insertReplacementText":
    case "insertCompositionText":
      return { kind: "typeOver", text: data ?? "" };
    case "insertParagraph":
      return { kind: "splitAtSelection" };
    case "insertLineBreak":
    case "deleteContentBackward":
    case "deleteContentForward":
    case "deleteWordBackward":
    case "deleteWordForward":
    case "deleteByCut":
      return { kind: "deleteSelection" };
    case "insertFromPaste":
    case "insertFromPasteAsQuotation":
      return { kind: "paste" };
    default:
      return { kind: "native" };
  }
}
