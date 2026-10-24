import {expect, test} from '@playwright/test';
import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import {readZipEntry} from './zip.js';

const root = new URL('../../TestFiles/RP/MoveRangeMembership/', import.meta.url);
const cases = JSON.parse(readFileSync(new URL('word-endpoints.json', root), 'utf8')).cases as {
  id: string; word: Record<string, {paragraphs: string[]}>;
}[];

for (const sample of cases) for (const worker of [false, true]) {
  test(`${sample.id}: ${worker ? 'worker lists one supported move' : 'main thread matches Word'}`, async ({page}) => {
    await page.goto(worker ? '/worker-test-harness.html' : '/test-harness.html');
    if (!worker) await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const bytes = Array.from(readFileSync(fileURLToPath(new URL(sample.id + '.docx', root))));
    const results = await page.evaluate(async ({bytes, worker}) => {
      if (worker) await (window as any).createDocxodusWorker();
      const api = worker ? (window as any).DocxodusWorker : (window as any).Docxodus;
      const results: {accept: boolean; route: string; bytes: number[]}[] = [];
      if (worker) {
        try {
          const revisions = await api.getRevisions(new Uint8Array(bytes));
          if (revisions.length !== 1 || revisions[0].type !== 'move'
              || revisions[0].resolutionStatus !== 'supported') throw new Error(JSON.stringify(revisions));
          if (typeof (window as any).Docxodus !== 'undefined') throw new Error('Unexpected main-thread runtime');
          return results;
        } finally {api.terminate();}
      }
      try {
        for (const accept of [false, true]) {
          if (!worker) {
            const output = accept ? api.DocxDiffBridge.AcceptRevisions(new Uint8Array(bytes))
              : api.DocxDiffBridge.RejectRevisions(new Uint8Array(bytes));
            results.push({accept, route: 'stateless', bytes: Array.from(output)});
          }
          for (const selective of [false, true]) {
            const session = (window as any).DocxodusSession.openDocxSession(new Uint8Array(bytes), api);
            try {
              const revisions = await session.listRevisions();
              if (revisions.length !== 1 || revisions[0].type !== 'move'
                  || revisions[0].resolutionStatus !== 'supported') throw new Error(JSON.stringify(revisions));
              const result = selective
                ? accept ? await session.acceptRevision(revisions[0].id) : await session.rejectRevision(revisions[0].id)
                : accept ? await session.acceptAllRevisions() : await session.rejectAllRevisions();
              if (!result.success) throw new Error(JSON.stringify(result));
              if ((await session.listRevisions()).length) throw new Error('Unresolved move');
              results.push({accept, route: selective ? 'selective' : 'bulk', bytes: Array.from(await session.save())});
            } finally {await session.close();}
          }
        }
      } finally {if (worker) api.terminate();}
      return results;
    }, {bytes, worker});
    expect(results).toHaveLength(worker ? 0 : 6);
    for (const result of results) {
      const xml = readZipEntry(Buffer.from(result.bytes), 'word/document.xml');
      const paragraphs = await page.evaluate(xml => {
        const document = new DOMParser().parseFromString(xml, 'application/xml');
        const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const body = document.getElementsByTagNameNS(namespace, 'body')[0];
        return Array.from(body.children).filter(element => element.localName === 'p')
          .map(p => Array.from(p.getElementsByTagNameNS(namespace, 't')).map(t => t.textContent).join(''));
      }, xml);
      expect(paragraphs, result.route).toEqual(sample.word[result.accept ? 'accept-all' : 'reject-all'].paragraphs);
      expect(xml).not.toMatch(/<w:move(?:From|To)(?:RangeStart|RangeEnd)?[\s/>]/);
      await test.info().attach(`${result.route}-${result.accept ? 'accept' : 'reject'}.docx`, {
        body: Buffer.from(result.bytes), contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    }
  });
}
