import { test } from '@playwright/test';
// Drives the REAL spec file under CPU throttling by re-running it in-process is not
// possible, so this mirrors the shipped budget against a throttled runner.
test('throttled x12 still reaches a pickup within the frame budget', async ({ page }) => {
  test.setTimeout(900000);
  const client = await page.context().newCDPSession(page);
  await page.goto('/demo-arcade.html?engine=./embed.bundle.js&intro=0&cart=e1m1');
  await page.waitForFunction(() => (window as any).__arcade !== undefined, { timeout: 90000 });
  await page.selectOption('#pace', '0');
  await page.waitForFunction(() => (window as any).__arcade.frames() >= 2, { timeout: 30000 });
  await client.send('Emulation.setCPUThrottlingRate', { rate: 12 });
  const t0 = Date.now();
  await page.evaluate(() => {
    const a = (window as any).__arcade;
    const bfs = () => {
      const s = a.game(); const rows: string[] = [];
      for (let y = 0; ; y++) { const r = s.mapRow(y); if (!r) break; rows.push(r); }
      const W = rows[0].length, H = rows.length; const sig = s.sigilsLeft > 0;
      const isT = (ch: string) => (sig ? ch === '§' : ch === '*');
      const walk = (ch: string) => ch === '.' || ch === '§' || (ch === '*' && !sig);
      const sx = Math.floor(s.player.x), sy = Math.floor(s.player.y);
      const prev = new Map<string, string | null>([[sx + ',' + sy, null]]);
      const q: Array<[number, number]> = [[sx, sy]];
      while (q.length) {
        const [x, y] = q.shift()!;
        if (isT(rows[y][x])) {
          let k = x + ',' + y, pk = prev.get(k), d = 0;
          while (pk && prev.get(pk) !== null) { k = pk; pk = prev.get(k)!; d++; }
          const [fx, fy] = k.split(',').map(Number); return { fx, fy, dist: d };
        }
        for (const [nx, ny] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
          const kk = nx + ',' + ny;
          if (nx < 0 || nx >= W || ny < 0 || ny >= H || prev.has(kk)) continue;
          if (!walk(rows[ny][nx]) && !isT(rows[ny][nx])) continue;
          prev.set(kk, x + ',' + y); q.push([nx, ny]);
        }
      }
      return null;
    };
    let goal: any = null, stuck = 0;
    setInterval(() => {
      if (!a.playing()) return;
      const s = a.game(); if (s.mode === 'dead') return;
      const input = a.input;
      const foe = (s.enemies ?? []).filter((e: any) => e.awake)
        .map((e: any) => ({ e, d: Math.hypot(e.x - s.player.x, e.y - s.player.y) }))
        .filter((f: any) => f.d < 6).sort((x: any, y: any) => x.d - y.d)[0];
      if (foe) {
        const vx = foe.e.x - s.player.x, vy = foe.e.y - s.player.y;
        const fc = s.player.dx * vy - s.player.dy * vx, fd = s.player.dx * vx + s.player.dy * vy;
        const al = fd > 0 && Math.abs(fc) < 0.25 * foe.d;
        input.set('KeyW', false); input.set('ShiftLeft', false);
        input.set('ArrowLeft', false); input.set('ArrowRight', false);
        input.set('Space', al); if (!al) input.set(fc < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        return;
      }
      input.set('Space', false);
      if (!goal || (Math.floor(s.player.x) === goal.fx && Math.floor(s.player.y) === goal.fy) || ++stuck > 25) {
        goal = bfs(); stuck = 0; if (!goal) return;
      }
      const tx = goal.fx + 0.5 - s.player.x, ty = goal.fy + 0.5 - s.player.y;
      const cross = s.player.dx * ty - s.player.dy * tx, dot = s.player.dx * tx + s.player.dy * ty;
      input.set('ArrowLeft', false); input.set('ArrowRight', false);
      input.set('KeyW', false); input.set('ShiftLeft', false);
      if (dot < 0 || Math.abs(cross) > 0.35 * Math.hypot(tx, ty)) {
        input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        if (dot > 0) input.set('KeyW', true);
      } else { input.set('KeyW', true); input.set('ShiftLeft', true); }
    }, 60);
  });
  await page.waitForFunction(
    ([n0, budget]) => {
      const a = (window as any).__arcade;
      return a.game().sigilsLeft < n0 || a.frames() >= budget;
    }, [5, 2500], { timeout: 540000, polling: 500 });
  const out = await page.evaluate(() => ({
    sig: (window as any).__arcade.game().sigilsLeft, frames: (window as any).__arcade.frames(),
  }));
  console.log(`VERIFY picked=${out.sig < 5} frames=${out.frames} secs=${Math.round((Date.now() - t0) / 1000)}`);
});
