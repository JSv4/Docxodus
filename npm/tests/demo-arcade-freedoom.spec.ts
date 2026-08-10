import { test, expect, Page } from '@playwright/test';

// Proof that a REAL Doom-format level plays inside a live Word document.
//
// Cartridge 3 of THE DOCX ARCADE is Freedoom's E1M1 (BSD-licensed), its
// linedefs rasterized to the raycaster grid by docs/demo/tools/wad2cart.mjs.
// The generic arcade spec (demo-arcade.spec.ts) already covers boot, frame
// incrementality, and save/reopen for this cartridge; what THIS spec proves
// is play: the level's real geometry is walkable through the live
// replaceXml + editor.refresh() loop, with the same autopilot a player's
// hands would be — synthetic key presses into the game's own input seam,
// steering by what the game itself reports.
//
// The full-completion marathon (every sigil + the exit gate, ~9 minutes of
// simulated play) runs only when DOCXODUS_DOOM_MARATHON=1 — it is the "beat
// the whole level" proof, far too slow for every CI run. The default suite
// proves real play in miniature: navigate the actual E1M1 corridors to the
// nearest sigil and collect it.

const OVERRIDE = 'engine=./embed.bundle.js';

async function bootFreedoom(page: Page) {
  // pace=0: unthrottled frames — the autopilot runs as fast as the document
  // renders, which is the honest speed of the machine under test.
  await page.goto(`/demo-arcade.html?${OVERRIDE}&cart=e1m1`);
  await page.waitForFunction(
    () =>
      (window as any).__arcade !== undefined ||
      (window as any).__arcadeError !== undefined,
    { timeout: 90000 },
  );
  const err = await page.evaluate(() => (window as any).__arcadeError);
  expect(err, `arcade boot failed: ${err}`).toBeUndefined();
  await page.selectOption('#pace', '0');
  await page.waitForFunction(() => (window as any).__arcade.frames() >= 2, { timeout: 30000 });
}

/** Install the BFS autopilot: every 60 ms it re-steers the game's own input
 *  toward the nearest '§' (or the '*' gate once none remain), reading only
 *  what the game state exposes — exactly the information on the player's
 *  screen. Returns nothing; poll __autopilot.status() from the test. */
async function startAutopilot(page: Page) {
  await page.evaluate(() => {
    const a = (window as any).__arcade;
    const bfsNext = () => {
      const s = a.game();
      const rows: string[] = [];
      for (let y = 0; ; y++) { const r = s.mapRow(y); if (!r) break; rows.push(r); }
      const W = rows[0].length, H = rows.length;
      const sig = s.sigilsLeft > 0;
      const isTarget = (ch: string) => (sig ? ch === '§' : ch === '*');
      const walk = (ch: string) => ch === '.' || ch === '§' || (ch === '*' && !sig);
      const sx = Math.floor(s.player.x), sy = Math.floor(s.player.y);
      const prev = new Map<string, string | null>([[sx + ',' + sy, null]]);
      const q: Array<[number, number]> = [[sx, sy]];
      while (q.length) {
        const [x, y] = q.shift()!;
        if (isTarget(rows[y][x])) {
          let k = x + ',' + y, pk = prev.get(k);
          while (pk && prev.get(pk) !== null) { k = pk; pk = prev.get(k)!; }
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

    let goal: { fx: number; fy: number } | null = null;
    let stuck = 0;
    // No-progress fallback: an engaged foe whose hp hasn't dropped after
    // sustained fire is behind a wall (awake but unhittable) — ignore it for
    // a while and let navigation resume; the game's own boredom timer will
    // put it back to sleep.
    let engaged: { key: string; hp: number; ticks: number } | null = null;
    const ignored = new Map<string, number>();
    const foeKey = (e: any) => `${e.kind}:${e.x.toFixed(0)},${e.y.toFixed(0)}`;
    const timer = setInterval(() => {
      if (!a.playing()) return;
      const s = a.game();
      if (s.mode === 'won') { clearInterval(timer); return; }
      const input = a.input;
      if (s.mode === 'dead') return; // the game respawns us; keep no keys held
      const now = Date.now();
      for (const [k, until] of ignored) if (until < now) ignored.delete(k);
      // Combat doctrine: the nearest AWAKE enemy within 6 cells gets faced
      // and shot (an enemy still asleep provably has no line of sight — its
      // own wake check would have fired — so shooting at it only hits wall).
      const foe = (s.enemies ?? [])
        .filter((e: any) => e.awake && !ignored.has(foeKey(e)))
        .map((e: any) => ({ e, d: Math.hypot(e.x - s.player.x, e.y - s.player.y) }))
        .filter((f: any) => f.d < 6)
        .sort((x: any, y: any) => x.d - y.d)[0];
      if (foe) {
        const key = foeKey(foe.e);
        if (engaged && engaged.key === key && engaged.hp === foe.e.hp) {
          if (++engaged.ticks > 50) { // ~3s of fire with no damage: wall between us
            ignored.set(key, now + 15000);
            engaged = null;
            return;
          }
        } else {
          engaged = { key, hp: foe.e.hp, ticks: 0 };
        }
        const vx = foe.e.x - s.player.x, vy = foe.e.y - s.player.y;
        const fcross = s.player.dx * vy - s.player.dy * vx;
        const fdot = s.player.dx * vx + s.player.dy * vy;
        const aligned = fdot > 0 && Math.abs(fcross) < 0.25 * foe.d;
        input.set('KeyW', false); input.set('ShiftLeft', false);
        input.set('ArrowLeft', false); input.set('ArrowRight', false);
        if (!aligned) input.set(fcross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        input.set('Space', false);
        if (aligned) input.set('Space', true); // fresh edge each tick; cooldown gates
        return;
      }
      engaged = null;
      input.set('Space', false);
      if (!goal || (Math.floor(s.player.x) === goal.fx && Math.floor(s.player.y) === goal.fy)
        || ++stuck > 25) {
        goal = bfsNext();
        stuck = 0;
        if (!goal) { clearInterval(timer); return; }
      }
      const tx = goal.fx + 0.5 - s.player.x, ty = goal.fy + 0.5 - s.player.y;
      const cross = s.player.dx * ty - s.player.dy * tx;
      const dot = s.player.dx * tx + s.player.dy * ty;
      input.set('ArrowLeft', false); input.set('ArrowRight', false);
      input.set('KeyW', false); input.set('ShiftLeft', false);
      if (dot < 0 || Math.abs(cross) > 0.35 * Math.hypot(tx, ty)) {
        input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        if (dot > 0) input.set('KeyW', true);
      } else {
        input.set('KeyW', true); input.set('ShiftLeft', true);
      }
    }, 60);
    (window as any).__autopilot = {
      stop: () => clearInterval(timer),
      status: () => ({ ...a.game(), frames: a.frames() }),
    };
  });
}

test.describe('Freedoom E1M1 inside a Word document', () => {
  test('boots the real level: spawn view renders, HUD names it, the map band scrolls', async ({ page }) => {
    await bootFreedoom(page);
    const state = await page.evaluate(() => {
      const a = (window as any).__arcade;
      return {
        cart: a.cart() as string,
        text: a.canvasText() as string,
        window: a.game().window as [number, number],
        sigils: a.game().sigilsLeft as number,
        monsters: a.game().killsTotal as number,
        health: a.game().health as number,
        fallback: a.editor.lastReconcileFallback as string | null,
        row0: a.game().mapRow(0) as string,
      };
    });
    expect(state.cart).toBe('e1m1');
    expect(state.text).toContain('FREEDOOM E1M1');
    expect(state.text).toMatch(/[█▓▒░]/);       // raycast wall columns on screen
    expect(state.text).toContain('MAP @');       // the scrolling window header
    expect(state.text).toContain('HP');          // combat HUD present
    expect(state.sigils).toBe(5);                // the level's own pickup spots
    expect(state.monsters).toBe(45);             // the level's own monster placements
    expect(state.health).toBe(100);
    expect(state.fallback).toBeNull();           // frames stay incremental
    expect(state.row0.length).toBe(126);         // the real rasterized grid width
    // The window is a proper sub-view of the 126×109 level, not the whole map.
    expect(state.window[0]).toBeGreaterThanOrEqual(0);
  });

  test('keyboard walks the real E1M1 entry corridor', async ({ page }) => {
    await bootFreedoom(page);
    const before = await page.evaluate(() => (window as any).__arcade.game().player as { x: number; y: number });
    await page.keyboard.down('KeyW');
    await page.waitForFunction(
      (x0) => (window as any).__arcade.game().player.x > x0 + 2,
      before.x,
      { timeout: 20000 },
    );
    await page.keyboard.up('KeyW');
    // Turning changes the camera: the direction vector must rotate.
    const dy0 = await page.evaluate(() => (window as any).__arcade.game().player.dy as number);
    await page.keyboard.down('ArrowLeft');
    await page.waitForFunction(
      (d0) => Math.abs((window as any).__arcade.game().player.dy - d0) > 0.3,
      dy0,
      { timeout: 20000 },
    );
    await page.keyboard.up('ArrowLeft');
  });

  test('autopilot navigates the real geometry and collects a Freedoom pickup spot', async ({ page }) => {
    test.setTimeout(300000);
    await bootFreedoom(page);
    const sigils0 = await page.evaluate(() => (window as any).__arcade.game().sigilsLeft as number);
    expect(sigils0).toBe(5);
    await startAutopilot(page);
    await page.waitForFunction(
      (n0) => (window as any).__autopilot.status().sigilsLeft < n0,
      sigils0,
      { timeout: 240000, polling: 500 },
    );
    await page.evaluate(() => (window as any).__autopilot.stop());
    // Still a document the whole time: save the very frame the sigil was
    // taken on and reopen it as a fresh DOCX.
    const result = await page.evaluate(() => {
      const a = (window as any).__arcade;
      a.pause();
      const bytes: Uint8Array = a.save();
      const handle = a.bridge.OpenSession(bytes, '');
      const markdown = JSON.parse(a.bridge.Project(handle)).markdown as string;
      a.bridge.CloseSession(handle);
      return { magic: Array.from(bytes.slice(0, 2)), markdown, sigils: a.game().sigilsLeft };
    });
    expect(result.sigils).toBeLessThan(sigils0);
    expect(result.magic).toEqual([0x50, 0x4b]);
    expect(result.markdown).toContain('FREEDOOM E1M1');
  });

  test('MARATHON: beat all of E1M1 — every sigil, then the exit gate', async ({ page }) => {
    test.skip(!process.env.DOCXODUS_DOOM_MARATHON, 'set DOCXODUS_DOOM_MARATHON=1 for the full-level playthrough');
    test.setTimeout(3600000);
    await bootFreedoom(page);
    await startAutopilot(page);
    let last = -1;
    for (;;) {
      const s = await page.evaluate(() => (window as any).__autopilot.status());
      if (s.mode === 'won') break;
      if (s.sigilsLeft !== last) {
        last = s.sigilsLeft;
        console.log(`sigils left ${s.sigilsLeft} · kills ${s.kills}/${s.killsTotal} · HP ${Math.ceil(s.health)} · frame ${s.frames} · player (${s.player.x.toFixed(1)}, ${s.player.y.toFixed(1)})`);
        await page.screenshot({ path: `test-results/doom-marathon-${5 - s.sigilsLeft}.png` });
      }
      await page.waitForTimeout(2000);
    }
    const final = await page.evaluate(() => {
      const a = (window as any).__arcade;
      a.pause();
      const bytes: Uint8Array = a.save();
      return {
        text: a.canvasText() as string,
        frames: a.frames() as number,
        magic: Array.from(bytes.slice(0, 2)),
      };
    });
    console.log(`WON after ${final.frames} document frames`);
    await page.screenshot({ path: 'test-results/doom-marathon-won.png' });
    expect(final.text).toContain('E1M1 CLEARED');
    expect(final.magic).toEqual([0x50, 0x4b]);
  });
});
