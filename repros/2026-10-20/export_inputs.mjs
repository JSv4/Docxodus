// Independent runner for synthetic examples, using the public export API.
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const [moduleFile, outputRoot, ...names] = process.argv.slice(2);
const { renderDocxFile } = await import(pathToFileURL(moduleFile));
await fs.mkdir(outputRoot, { recursive: true });
for (const name of names) {
  const input = path.resolve("inputs", `${name}.docx`);
  for (const presentation of (["footnote-format", "note-arial20"].includes(name) ? ["docxodus", "word"] : ["docxodus"])) {
    const prefix = path.join(outputRoot, `${name}-${presentation}`);
    const began = Date.now();
    try {
      const result = await renderDocxFile(input, {
        pdfPath: `${prefix}.pdf`,
        htmlPath: `${prefix}.html`,
        pageMapPath: `${prefix}.pages.json`,
        reportPath: `${prefix}.report.json`,
      }, {
        reviewProfile: "markup",
        commentProfile: "hidden",
        markupPresentation: presentation,
        timeoutMs: 60000,
        ...(process.env.DOCXODUS_CHROMIUM ? { browserExecutablePath: process.env.DOCXODUS_CHROMIUM } : {}),
      });
      const entry = { name, presentation, ok: true, elapsedMs: Date.now()-began };
      await fs.writeFile(`${prefix}.status.json`, JSON.stringify(entry, null, 2));
      console.log(JSON.stringify(entry));
    } catch (error) {
      const entry = { name, presentation, ok: false, elapsedMs: Date.now()-began,
        message: String(error), code: error.code, stage: error.stage,
        details: error.details };
      await fs.writeFile(`${prefix}.status.json`, JSON.stringify(entry, null, 2));
      console.log(JSON.stringify(entry));
    }
  }
}
