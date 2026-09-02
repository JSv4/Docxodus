import { test } from '@playwright/test';
const OVERRIDE = 'engine=./embed.bundle.js&intro=0';

for (const rate of [1, 6, 12]) {
  test(`throttle x${rate}: autopilot progress`, async ({ page }) => {
    test.setTimeout(400000);
    const client = await page.context().newCDPSession(page);
    await page.goto(`/demo-arcade.html?${OVERRIDE}&cart=e1m1`);
    await page.waitForFunction(() => (window as any).__arcade !== undefined, { timeout: 90000 });
    await page.selectOption('#pace', '0');
    await page.waitForFunction(() => (window as any).__arcade.frames() >= 2, { timeout: 30000 });
    if (rate > 1) await client.send('Emulation.setCPUThrottlingRate', { rate });

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
      const diag: any = { ticks: 0, bfsCalls: 0, bfsNull: 0, minDist: Infinity, combat: 0 };
      (window as any).__d = diag;
      const timer = setInterval(() => {
        diag.ticks++;
        if (!a.playing()) return;
        const s = a.game();
        if (s.mode === 'won') { clearInterval(timer); return; }
        if (s.mode === 'dead') return;
        const input = a.input;
        const foe = (s.enemies ?? []).filter((e: any) => e.awake)
          .map((e: any) => ({ e, d: Math.hypot(e.x - s.player.x, e.y - s.player.y) }))
          .filter((f: any) => f.d < 6).sort((x: any, y: any) => x.d - y.d)[0];
        if (foe) {
          diag.combat++;
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
          diag.bfsCalls++; goal = bfs(); stuck = 0;
          if (!goal) { diag.bfsNull++; return; }
          diag.minDist = Math.min(diag.minDist, goal.dist);
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

    const t0 = Date.now();
    let picked = false;
    try {
      await page.waitForFunction(() => (window as any).__arcade.game().sigilsLeft < 5, undefined,
        { timeout: 240000, polling: 500 });
      picked = true;
    } catch { /* record below */ }
    const out = await page.evaluate(() => ({
      d: (window as any).__d, frames: (window as any).__arcade.frames(),
      player: (window as any).__arcade.game().player, sig: (window as any).__arcade.game().sigilsLeft,
      health: (window as any).__arcade.game().health,
    }));
    console.log(`THROTTLE=${rate} picked=${picked} secs=${Math.round((Date.now() - t0) / 1000)} ` + JSON.stringify(out));
  });
}
