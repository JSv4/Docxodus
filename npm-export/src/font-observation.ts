import { createHash } from "node:crypto";
import type { BrowserContext, Page } from "playwright-core";
import { FONT_REQUEST_ATTRIBUTE, type CompleteRenderReport, type FontObservation } from "docxodus/export-browser";

// Serialized into the script-disabled, closed snapshot realm. No DOM writes or character splitting.
function textNodes(maximum: number): Text[] {
  const result: Text[] = [];
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!node.nodeValue?.trim() || node.parentElement?.closest("head,script,style,template,noscript")) continue;
    result.push(node as Text);
    if (result.length > maximum) break;
  }
  return result;
}

function textIdentity(this: Text, requestAttribute: string) {
  const parent = this.parentElement;
  return {
    text: this.nodeValue ?? "",
    requestId: parent?.closest(`[${requestAttribute}]`)?.getAttribute(requestAttribute),
    anchorId: parent?.closest("[data-source-anchor-id]")?.getAttribute("data-source-anchor-id"),
    fragmentId: parent?.closest("[data-page-fragment-id]")?.getAttribute("data-page-fragment-id"),
    pageNumber: Number(parent?.closest("[data-page-number]")?.getAttribute("data-page-number")),
  };
}

/** CDP reports shaping usage for individual text nodes, including SVG, without font-file identity. */
export async function observeFonts(
  context: BrowserContext,
  page: Page,
  report: CompleteRenderReport,
  media: "screen" | "print",
): Promise<FontObservation> {
  const remaining = report.options.policy.limits.renderDiagnostics - report.fonts.length
    - report.resources.length - report.warnings.length - report.unsupportedContent.length;
  const requests = new Set(report.fonts.map(font => font.requestId));
  let session;
  try {
    session = await context.newCDPSession(page);
    await session.send("DOM.enable");
    await session.send("CSS.enable");
    await session.send("DOM.getDocument", { depth: 0 });
    const array = await session.send("Runtime.evaluate", {
      expression: `(${textNodes.toString()})(${Math.max(0, remaining)})`,
      objectGroup: "docxodus-font-observation",
    });
    if (!array.result.objectId || array.exceptionDetails) throw new Error("Cannot enumerate finalized text nodes.");
    const properties = await session.send("Runtime.getProperties", { objectId: array.result.objectId, ownProperties: true });
    const nodes = properties.result.filter(property => /^(?:0|[1-9][0-9]*)$/.test(property.name));
    if (nodes.length > remaining) return { status: "unavailable", reason: "diagnostic_limit" };
    const uses: Extract<FontObservation, { status: "observed" }>["uses"] = [];
    let admitted = 0;
    for (const property of nodes) {
      const objectId = property.value?.objectId;
      if (!objectId) throw new Error("Finalized text node has no protocol identity.");
      const { nodeId } = await session.send("DOM.requestNode", { objectId });
      const { fonts } = await session.send("CSS.getPlatformFontsForNode", { nodeId });
      if (fonts.length === 0) continue;
      if (admitted + 1 + fonts.length > remaining) return { status: "unavailable", reason: "diagnostic_limit" };
      admitted += 1 + fonts.length;
      const validName = (value: unknown): value is string => typeof value === "string" && value.length > 0
        && value.length <= 2048 && !/[\u0000-\u001f\u007f]/.test(value);
      if (fonts.some(font => !validName(font.familyName) || typeof font.isCustomFont !== "boolean"
        || !Number.isSafeInteger(font.glyphCount) || font.glyphCount <= 0
        || (font.postScriptName !== "" && !validName(font.postScriptName)))) {
        throw new Error("Chromium returned invalid platform-font metadata.");
      }
      const identity = await session.send("Runtime.callFunctionOn", {
        objectId, functionDeclaration: textIdentity.toString(), arguments: [{ value: FONT_REQUEST_ATTRIBUTE }], returnByValue: true,
      });
      const value = identity.result.value as ReturnType<typeof textIdentity> | undefined;
      if (!value || identity.exceptionDetails || typeof value.text !== "string") throw new Error("Cannot read finalized text identity.");
      uses.push({
        ...(value.requestId && requests.has(value.requestId) ? { requestId: value.requestId } : {}),
        ...(value.anchorId ? { anchorId: value.anchorId } : {}),
        ...(value.fragmentId ? { fragmentId: value.fragmentId } : {}),
        ...(Number.isSafeInteger(value.pageNumber) && value.pageNumber > 0 ? { pageNumber: value.pageNumber } : {}),
        textDigest: createHash("sha256").update(value.text).digest("hex"),
        textCodePointCount: Array.from(value.text).length,
        faces: fonts.map(font => ({ family: font.familyName,
          ...(font.postScriptName ? { postscriptName: font.postScriptName } : {}),
          isCustomFont: font.isCustomFont, glyphCount: font.glyphCount,
        })).sort((a, b) => {
          const left = JSON.stringify(a), right = JSON.stringify(b);
          return left < right ? -1 : left > right ? 1 : 0;
        }),
      });
    }
    return { status: "observed", method: "chromium-cdp", media, fileIdentity: "unavailable", uses };
  } catch (error) {
    if (error instanceof Error && /(?:CSS\.(?:enable|getPlatformFontsForNode).*wasn't found|CDP sessions are only available in Chromium)/.test(error.message))
      return { status: "unavailable", reason: "protocol_unavailable" };
    throw error;
  } finally {
    if (session) {
      await session.send("Runtime.releaseObjectGroup", { objectGroup: "docxodus-font-observation" }).catch(() => undefined);
      await session.detach().catch(() => undefined);
    }
  }
}
