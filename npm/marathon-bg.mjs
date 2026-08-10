import { chromium } from '@playwright/test';
import { writeFileSync } from 'node:fs';
const OUT = '/tmp/claude-0/-home-user-Docxodus/b28cb601-65ad-57dd-925a-334c061082c3/scratchpad/shots';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 950 }, deviceScaleFactor: 2 });
await page.addInitScript(() => {
  const real = performance.now.bind(performance);
  const t0 = real();
  performance.now = () => t0 + (real() - t0) * 3;
});
await page.goto('http://localhost:8082/demo-arcade.html?engine=./embed.bundle.js&cart=e1m1');
await page.waitForFunction(() => window.__arcade !== undefined || window.__arcadeError !== undefined, null, { timeout: 120000 });
const err = await page.evaluate(() => window.__arcadeError);
if (err) { console.log('BOOT FAILED:', err); process.exit(1); }
await page.waitForFunction(() => window.__arcade.frames() >= 3, null, { timeout: 60000 });
await page.selectOption('#pace', '0');
await page.evaluate(() => {
  const a = window.__arcade;
  const bfsNext = () => {
    const s = a.game();
    const rows = [];
    for (let y = 0; ; y++) { const r = s.mapRow(y); if (!r) break; rows.push(r); }
    const W = rows[0].length, H = rows.length;
    const sig = s.sigilsLeft > 0;
    const isTarget = (ch) => (sig ? ch === '§' : ch === '*');
    const walk = (ch) => ch === '.' || ch === '§' || (ch === '*' && !sig);
    const sx = Math.floor(s.player.x), sy = Math.floor(s.player.y);
    const prev = new Map([[sx + ',' + sy, null]]);
    const q = [[sx, sy]];
    while (q.length) {
      const [x, y] = q.shift();
      if (isTarget(rows[y][x])) {
        let k = x + ',' + y, pk = prev.get(k);
        while (pk && prev.get(pk) !== null) { k = pk; pk = prev.get(k); }
        const [fx, fy] = k.split(',').map(Number);
        return { fx, fy };
      }
      for (const [nx, ny] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
        const kk = nx + ',' + ny;
        if (nx < 0 || nx >= W || ny < 0 || ny >= H || prev.has(kk)) continue;
        if (!walk(rows[ny][nx]) && !isTarget(rows[ny][nx])) continue;
        prev.set(kk, x + ',' + y);
        q.push([nx, ny]);
      }
    }
    return null;
  };
  let goal = null, stuck = 0;
  window.__pilot = setInterval(() => {
    const s = a.game();
    if (s.mode !== 'run') return;
    const input = a.input;
    const foe = (s.enemies ?? [])
      .filter((e) => e.awake)
      .map((e) => ({ e, d: Math.hypot(e.x - s.player.x, e.y - s.player.y) }))
      .filter((f) => f.d < 6)
      .sort((x, y) => x.d - y.d)[0];
    if (foe) {
      const vx = foe.e.x - s.player.x, vy = foe.e.y - s.player.y;
      const cross = s.player.dx * vy - s.player.dy * vx;
      const dot = s.player.dx * vx + s.player.dy * vy;
      const aligned = dot > 0 && Math.abs(cross) < 0.25 * foe.d;
      input.set('KeyW', false); input.set('ShiftLeft', false);
      input.set('ArrowLeft', false); input.set('ArrowRight', false);
      if (!aligned) input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
      input.set('Space', false);
      if (aligned) input.set('Space', true);
      return;
    }
    input.set('Space', false);
    if (!goal || (Math.floor(s.player.x) === goal.fx && Math.floor(s.player.y) === goal.fy) || ++stuck > 25) {
      goal = bfsNext(); stuck = 0;
      if (!goal) return;
    }
    const tx = goal.fx + 0.5 - s.player.x, ty = goal.fy + 0.5 - s.player.y;
    const cross = s.player.dx * ty - s.player.dy * tx;
    const dot = s.player.dx * tx + s.player.dy * ty;
    input.set('ArrowLeft', false); input.set('ArrowRight', false);
    input.set('KeyW', false); input.set('ShiftLeft', false);
    if (dot < 0 || Math.abs(cross) > 0.35 * Math.hypot(tx, ty)) {
      input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
      if (dot > 0) input.set('KeyW', true);
    } else { input.set('KeyW', true); input.set('ShiftLeft', true); }
  }, 40);
});
let last = -1;
let polls = 0;
const t0 = Date.now();
for (;;) {
  if (Date.now() - t0 > 1800000) { console.log('TIMEOUT'); break; }
  const s = await page.evaluate(() => {
    const a = window.__arcade;
    const g = a.game();
    return { mode: g.mode, sig: g.sigilsLeft, kills: g.kills, hp: Math.ceil(g.health), frames: a.frames(),
      px: g.player.x, py: g.player.y };
  });
  if (s.sig !== last || ++polls % 40 === 0) {
    last = s.sig;
    console.log(`sigils left ${s.sig} · kills ${s.kills}/45 · HP ${s.hp} · frame ${s.frames} · (${s.px.toFixed(1)},${s.py.toFixed(1)}) · mode ${s.mode} · ${((Date.now() - t0) / 1000).toFixed(0)}s`);
  }
  if (s.mode === 'won') {
    console.log(`WON after ${s.frames} document frames · kills ${s.kills}/45 · ${((Date.now() - t0) / 1000).toFixed(0)}s wall`);
    await page.evaluate(() => window.clearInterval(window.__pilot));
    await page.screenshot({ path: `${OUT}/11-e1m1-cleared.png` });
    const saved = await page.evaluate(() => {
      const a = window.__arcade;
      a.pause();
      const bytes = a.save();
      return Array.from(bytes);
    });
    writeFileSync(`${OUT}/e1m1-cleared-frame.docx`, Buffer.from(saved));
    console.log('saved winning frame:', saved.length, 'bytes, PK =', (saved[0] === 0x50 && saved[1] === 0x4b));
    break;
  }
  await page.waitForTimeout(1500);
}
await browser.close();
console.log('MARATHON-DONE');
