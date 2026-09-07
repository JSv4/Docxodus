// THE DOCX ARCADE — its controls, in one place.
//
// Two pages host the arcade now: `arcade.html` (the cabinet, full screen) and
// the landing page, which mounts the arcade instead of the plain editor on a
// phone. Both need the same controls, so the controls live here rather than
// being hand-written twice — the same reason the editor surface itself ships
// from `npm/src/ribbon.ts` instead of once per demo page, after three
// hand-rolled copies of it drifted.
//
// Density is measured from the CONTROLS' OWN HOST, not from a viewport media
// query — the landing page frames the arcade inside a card, and a narrow card
// in a wide page is narrow. That is the same rule `mountRibbon` applies to the
// ribbon it draws directly above these controls.
//
//   wide    one bar under the document: cartridges, transport, pacing, embed,
//           telemetry, hint. Unchanged from the cabinet's original dock.
//   compact a slim HUD strip keeps the two controls you touch mid-game
//           (play/pause and pacing); cartridges, restart, embed, telemetry and
//           the hint move behind a "⋯" sheet. A thumb D-pad and an action
//           cluster float over the bottom corners of the game — where the
//           thumbs already are, and clear of the centre of the screen they are
//           steering. Nothing is dropped, only re-placed.
//
// WHAT THE PAD OFFERS IS THE CARTRIDGE'S CALL, not this module's. `setPad`
// takes a touch profile — fixed slots, because the geometry is fixed — and a
// cartridge fills in only the slots it has a key for: the platformer wants
// three buttons, the raycasters want strafe and a run latch, and Doom wants
// USE, its own menu, the automap and seven weapons as well. A phone that can
// only walk and shoot cannot open a door, and E1M1's first door is thirty
// seconds in, which is where this started.
//
// The pad is deliberately NOT a descendant of the editor root. The driver
// pauses on any pointerdown inside the document — "the frame you clicked is
// now an ordinary paragraph with your caret in it" — so a control living in
// there would pause the game on every tap. It overlays the surface as a
// sibling instead.
//
// Compact layout moves the nodes themselves rather than duplicating them into
// two hidden layouts: one `#playpause`, one `#pace`, one set of cartridge
// buttons, whichever way the host is sized. The driver (`startArcade`) wires
// its listeners to those nodes once and never learns that layout exists.

const BREAKPOINT = 640;

/** What the pad offers before any cartridge has spoken for it — the arcade's
 *  common denominator, and what a host that never calls `setPad` keeps. */
const DEFAULT_PAD = {
  up: { code: 'KeyW', glyph: '▲', label: 'Forward / jump' },
  down: { code: 'KeyS', glyph: '▼', label: 'Back' },
  left: { code: 'ArrowLeft', glyph: '◀', label: 'Left / turn left' },
  right: { code: 'ArrowRight', glyph: '▶', label: 'Right / turn right' },
  fire: { code: 'Space', glyph: 'FIRE', label: 'Fire / jump / start' },
};

const CSS = `
.dxa-controls { position: fixed; inset: 0; z-index: 60; pointer-events: none; }
.dxa-controls[data-anchor="host"] { position: absolute; }
.dxa-controls * { box-sizing: border-box; }

.dxa-dock {
  position: absolute; left: 50%; bottom: 14px; transform: translateX(-50%);
  max-width: min(980px, calc(100% - 24px));
  display: flex; flex-direction: column; gap: 5px; align-items: center;
  pointer-events: auto;
  background: rgba(11, 15, 26, 0.94); color: #d5dce8;
  border: 1px solid #1f2a3f; border-radius: 10px;
  padding: 8px 14px 7px;
  font: 12px/1.4 "SF Mono", "Cascadia Code", Consolas, "Courier New", monospace;
  box-shadow: 0 10px 30px rgba(0, 0, 0, 0.45);
}
/* Author display values would otherwise beat the UA's [hidden] rule. */
.dxa-dock[hidden], .dxa-pad[hidden] { display: none; }

.dxa-dock .row { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: center; }
.dxa-dock button, .dxa-dock select {
  appearance: none; background: #111827; color: #d5dce8;
  border: 1px solid #1f2a3f; border-radius: 6px;
  font: inherit; font-size: 11.5px; padding: 4px 10px; cursor: pointer;
}
.dxa-dock button:hover { background: #18233a; }
.dxa-dock button[aria-pressed="true"] { background: #123a34; color: #5eead4; }
.dxa-dock button:disabled { opacity: 0.4; cursor: default; }
.dxa-playpause { min-width: 110px; }
.dxa-stats { color: #7a8699; font-size: 11px; white-space: nowrap; overflow-x: auto; max-width: 100%; }
.dxa-stats b { color: #d5dce8; font-weight: 600; }
.dxa-stats .inc { color: #5eead4; }
.dxa-hint { color: #7a8699; font-size: 10.5px; text-align: center; }
.dxa-hint b { color: #d5dce8; font-weight: 600; }
.dxa-embed { border-color: #123a34; color: #5eead4; }
.dxa-embed:hover { background: #123a34; }

.dxa-sheet, .dxa-more, .dxa-pad { display: none; }

/* ── compact ───────────────────────────────────────────────────────── */
.dxa-controls[data-compact="true"] .dxa-dock {
  left: 10px; right: 10px; bottom: 10px; transform: none;
  max-width: none; align-items: stretch; padding: 6px 8px;
  /* Cleared by the ribbon's own bottom chrome and by the home-indicator strip. */
  bottom: max(10px, env(safe-area-inset-bottom));
}
.dxa-controls[data-compact="true"] .dxa-strip {
  flex-wrap: nowrap; justify-content: space-between; gap: 8px;
}
.dxa-controls[data-compact="true"] .dxa-playpause {
  min-width: 0; flex: 1 1 auto; overflow: hidden;
  text-overflow: ellipsis; white-space: nowrap;
}
.dxa-controls[data-compact="true"] .dxa-dock button,
.dxa-controls[data-compact="true"] .dxa-dock select {
  min-height: 34px; font-size: 12px;
}
.dxa-controls[data-compact="true"] .dxa-more { display: block; flex: 0 0 auto; min-width: 38px; }
.dxa-controls[data-compact="true"] .dxa-sheet[data-open="true"] {
  display: flex; flex-direction: column; gap: 7px; align-items: center;
  max-height: 44vh; overflow-y: auto;
  margin: 0 0 7px; padding: 0 0 7px;
  border-bottom: 1px solid #1f2a3f;
}
/* Telemetry is one long line; in the wide dock it scrolls sideways, but in the
   sheet there is a column to wrap into and nothing to scroll it with. */
.dxa-controls[data-compact="true"] .dxa-stats {
  white-space: normal; overflow-x: visible; text-align: center;
}

/* Thumb controls. The wrapper spans the host so the two clusters can sit in
   its bottom corners; only the clusters themselves take pointer events, so
   the document between them stays tappable (and tapping it still pauses). */
.dxa-controls[data-compact="true"] .dxa-pad { display: block; }
/* The pad sits a fixed distance above the dock, so an opened sheet — which
   grows the dock upward — would be covered by it, and the sheet's own controls
   would be unreachable under a D-pad. While the menu is open you are not
   steering anything, so the pad stands down. */
.dxa-controls[data-menu="open"] .dxa-pad { display: none; }
.dxa-pad {
  position: absolute; left: 0; right: 0;
  bottom: calc(58px + env(safe-area-inset-bottom));
  pointer-events: none;
}
.dxa-pad button {
  appearance: none; pointer-events: auto; touch-action: none; user-select: none;
  -webkit-user-select: none; -webkit-tap-highlight-color: transparent;
  display: grid; place-items: center;
  color: #d5dce8; background: rgba(11, 15, 26, 0.82);
  border: 1px solid #33415c; border-radius: 12px;
  font: 16px/1 system-ui, sans-serif; cursor: pointer;
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.4);
}
/* A slot the current cartridge does not use is hidden, and the author
   author display value above would otherwise beat the UA's [hidden] rule. */
.dxa-pad button[hidden] { display: none; }
.dxa-pad button:active { color: #5eead4; background: #123a34; border-color: #14b8a6; }
/* Latched modifiers (RUN) read as pressed until you tap them off. */
.dxa-pad button[aria-pressed="true"] { color: #5eead4; background: #123a34; border-color: #14b8a6; }

/* Movement, as a 3 × 3 whose corners the cartridge fills in only if it has
   something to put there: turning rotates (↺ ↻) on the middle row, sidestep
   translates (◀ ▶) on the bottom corners beside it, so the two pairs never
   read as the same control. The platformer has neither and uses the middle
   row to walk. */
.dxa-dpad {
  position: absolute; left: 12px; bottom: 0;
  display: grid; gap: 4px;
  grid-template-columns: repeat(3, 46px); grid-template-rows: repeat(3, 42px);
}
.dxa-dpad button { width: 46px; height: 42px; font-size: 19px; font-weight: 600; }
.dxa-up { grid-area: 1 / 2; }
.dxa-left { grid-area: 2 / 1; }
.dxa-right { grid-area: 2 / 3; }
.dxa-down { grid-area: 3 / 2; }
.dxa-strafe-left { grid-area: 3 / 1; }
.dxa-strafe-right { grid-area: 3 / 3; }

/* Actions, stacked under the right thumb: FIRE where the thumb rests, USE
   directly above it (Doom's other one-hand-on-the-phone verb), RUN beside it,
   and the tray toggle on top — the rows collapse for a cartridge that wants
   none of them. */
.dxa-actions {
  position: absolute; right: 12px; bottom: 0;
  display: grid; gap: 8px; justify-items: end; align-items: end;
  grid-template-areas: "keys keys" "run use" "fire fire";
}
.dxa-pad .dxa-keys {
  grid-area: keys; min-width: 58px; height: 34px;
  font-size: 11px; letter-spacing: .06em;
}
.dxa-pad .dxa-run { grid-area: run; width: 58px; height: 56px; font-size: 11px; font-weight: 700; }
.dxa-pad .dxa-use {
  grid-area: use; width: 58px; height: 56px; border-radius: 50%;
  font-size: 11px; font-weight: 700;
}
.dxa-pad .dxa-fire {
  grid-area: fire; width: 74px; height: 74px; border-radius: 50%;
  font-size: 11px; font-weight: 700; letter-spacing: .08em;
  color: #5eead4; border-color: #14b8a6;
}

/* The tray: the keys a Doom player reaches for between fights rather than
   during one — weapons, the automap, Doom's own menu and its Enter. It opens
   over the game above the action cluster (180px of buttons + its gap) and
   stays open, because working Doom's menu takes several taps in a row. */
.dxa-extras {
  position: absolute; right: 12px; left: 12px; bottom: 190px;
  display: none; flex-wrap: wrap; gap: 6px; justify-content: flex-end;
}
.dxa-extras[data-open="true"] { display: flex; }
.dxa-pad .dxa-extras button {
  min-width: 40px; height: 38px; padding: 0 8px;
  border-radius: 10px; font-size: 12px; font-weight: 600;
}

@media (prefers-reduced-motion: reduce) {
  .dxa-controls * { transition: none !important; }
}
`;

function injectStyle() {
  if (document.getElementById('dxa-controls-style')) return;
  const style = document.createElement('style');
  style.id = 'dxa-controls-style';
  style.textContent = CSS;
  document.head.append(style);
}

function el(tag, props = {}, children = []) {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

/**
 * Build the Arcade's controls inside `host` and return the `ui` object
 * `startArcade({ ui })` expects.
 *
 * `host`      element the controls are placed in. With `anchor: 'host'` it
 *             must be positioned (the controls overlay it); with the default
 *             `'viewport'` they are fixed and `host` is only their parent.
 * `anchor`    'viewport' (the cabinet) | 'host' (framed inside a card).
 * `embed`     optional iframe snippet; supplies the Embed button when given.
 * `ids`       false to skip the historical bare ids (`#dock`, `#playpause`,
 *             …). They are what the specs and the README address, so they are
 *             on by default and simply must not be minted twice on one page.
 */
export function mountArcadeDock(host, { anchor = 'viewport', embed = null, ids = true } = {}) {
  injectStyle();

  const id = (name) => (ids ? { id: name } : {});

  const carts = el('div', { className: 'row dxa-carts', ...id('dockcarts') });
  const playpause = el('button', {
    className: 'dxa-playpause', ...id('playpause'), textContent: '⏸ Pause & edit',
  });
  const restart = el('button', {
    className: 'dxa-restart', ...id('restart'), title: 'Restart the cartridge (R)',
    textContent: '↻ Restart',
  });
  const pace = el('select', { className: 'dxa-pace', ...id('pace'), title: 'Target frame pacing' }, [
    el('option', { value: '125', textContent: '8 fps' }),
    el('option', { value: '100', textContent: '10 fps', selected: true }),
    el('option', { value: '66', textContent: '15 fps' }),
    el('option', { value: '0', textContent: 'unthrottled' }),
  ]);
  const stats = el('span', { className: 'dxa-stats', ...id('dockstats'), textContent: 'warming up…' });
  const rendering = el('select', { ...id('rendering'), title: 'DOOM rendering mode', hidden: true }, [
    el('option', { value: 'image', textContent: 'DOOM · Original' }),
    el('option', { value: 'ascii', textContent: 'DOOM · ASCII' }),
  ]);
  rendering.setAttribute('aria-label', 'DOOM rendering mode');
  const hint = el('div', { className: 'dxa-hint', ...id('dockhint'), textContent: 'warming up…' });
  const embedButton = embed
    ? el('button', {
        className: 'dxa-embed', ...id('copyEmbed'),
        title: 'Copy an iframe that embeds this arcade on your site',
        innerHTML: '&lt;/&gt; Embed',
      })
    : null;

  const more = el('button', {
    className: 'dxa-more', ...id('dockmore'), type: 'button',
    title: 'Cartridges and more controls', textContent: '⋯',
  });
  more.setAttribute('aria-label', 'Cartridges and more controls');
  more.setAttribute('aria-expanded', 'false');

  const sheet = el('div', { className: 'dxa-sheet', ...id('dockmenu') });
  const strip = el('div', { className: 'row dxa-strip' });
  const dock = el('div', { className: 'dxa-dock', ...id('dock'), hidden: true }, [sheet, strip]);

  const padButton = (cls) => el('button', { className: cls, type: 'button' });
  const slots = {
    up: padButton('dxa-up'),
    down: padButton('dxa-down'),
    left: padButton('dxa-left'),
    right: padButton('dxa-right'),
    strafeLeft: padButton('dxa-strafe-left'),
    strafeRight: padButton('dxa-strafe-right'),
    // Space is jump in the platformer, fire in the raycasters and in Doom, and
    // the coin drop on the attract screen — the one button a phone was missing
    // entirely, which is why Freedoom could be walked but not fought on a
    // touch screen. USE is the one it was missing after that: doors.
    fire: padButton('dxa-fire'),
    use: padButton('dxa-use'),
    run: padButton('dxa-run'),
  };
  const dpad = el('div', { className: 'dxa-dpad' },
    [slots.up, slots.left, slots.right, slots.down, slots.strafeLeft, slots.strafeRight]);

  const extras = el('div', { className: 'dxa-extras', ...id('padextras') });
  extras.setAttribute('aria-label', 'More game keys');
  const keys = el('button', { className: 'dxa-keys', ...id('padkeys'), type: 'button' });
  keys.setAttribute('aria-label', 'More game keys');
  keys.setAttribute('aria-expanded', 'false');
  keys.textContent = 'KEYS';

  const actions = el('div', { className: 'dxa-actions' }, [keys, slots.run, slots.use, slots.fire]);
  const pad = el('div', { className: 'dxa-pad', ...id('pad') }, [extras, dpad, actions]);
  pad.setAttribute('aria-label', 'Touch controls');

  const setTray = (open) => {
    extras.dataset.open = String(open && extras.childElementCount > 0);
    keys.setAttribute('aria-expanded', extras.dataset.open);
  };
  keys.addEventListener('click', () => setTray(extras.dataset.open !== 'true'));
  setTray(false);

  /** Fill one pad slot from a profile entry, or hide it. A hidden slot keeps
   *  its grid cell — the pad must not re-flow under a thumb when a cartridge
   *  changes — but loses its `data-code`, so a stale key can never be sent. */
  const applySlot = (button, spec) => {
    if (!spec) {
      button.hidden = true;
      button.removeAttribute('data-code');
      button.removeAttribute('data-mode');
      button.removeAttribute('aria-pressed');
      return;
    }
    button.hidden = false;
    button.setAttribute('data-code', spec.code);
    button.textContent = spec.glyph;
    button.setAttribute('aria-label', spec.label);
    button.title = spec.label;
    if (spec.toggle) {
      button.dataset.mode = 'toggle';
      if (!button.hasAttribute('aria-pressed')) button.setAttribute('aria-pressed', 'false');
    } else {
      button.removeAttribute('data-mode');
      button.removeAttribute('aria-pressed');
    }
  };

  /** Point the pad at a cartridge's touch profile. Slots the profile omits are
   *  hidden; `extras` (empty for every cartridge but Doom) fills the tray and
   *  supplies — or withdraws — its toggle. */
  function setPad(profile = DEFAULT_PAD) {
    for (const name of Object.keys(slots)) applySlot(slots[name], profile[name] ?? null);
    extras.replaceChildren(...(profile.extras ?? []).map((spec) => {
      const chip = el('button', { type: 'button', textContent: spec.glyph });
      chip.setAttribute('data-code', spec.code);
      chip.setAttribute('aria-label', spec.label);
      chip.title = spec.label;
      return chip;
    }));
    keys.hidden = extras.childElementCount === 0;
    if (keys.hidden) setTray(false);
  }
  setPad();

  const controls = el('div', { className: 'dxa-controls' }, [dock, pad]);
  controls.dataset.anchor = anchor;
  controls.dataset.compact = 'false';
  host.append(controls);

  const setMenu = (open) => {
    sheet.dataset.open = String(open);
    controls.dataset.menu = open ? 'open' : 'closed';
    more.setAttribute('aria-expanded', String(open));
  };
  more.addEventListener('click', () => setMenu(sheet.dataset.open !== 'true'));
  setMenu(false);

  let compact = null;
  function layout(next) {
    if (next === compact) return;
    compact = next;
    controls.dataset.compact = String(next);
    if (next) {
      // Everything you are not touching mid-frame goes behind "⋯"; the strip
      // keeps transport and pacing, and gains the sheet toggle.
      sheet.append(carts, rendering, ...(embedButton ? [restart, embedButton] : [restart]), stats, hint);
      strip.append(playpause, pace, more);
    } else {
      setMenu(false);
      strip.append(playpause, restart, pace, rendering, ...(embedButton ? [embedButton] : []), stats);
      dock.append(carts, strip, hint); // the cabinet's original three rows
    }
  }

  const measured = anchor === 'viewport' ? document.documentElement : host;
  const coarse = window.matchMedia('(pointer: coarse)');
  const remeasure = () => layout(measured.clientWidth <= BREAKPOINT || coarse.matches);
  const observer = new ResizeObserver(remeasure);
  observer.observe(measured);
  coarse.addEventListener('change', remeasure);
  remeasure();

  if (embedButton) {
    embedButton.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(embed);
      } catch {
        const area = Object.assign(document.createElement('textarea'), { value: embed });
        area.style.cssText = 'position:fixed;opacity:0';
        document.body.append(area);
        area.select();
        document.execCommand('copy');
        area.remove();
      }
      const previous = embedButton.innerHTML;
      embedButton.textContent = 'Copied!';
      setTimeout(() => { embedButton.innerHTML = previous; }, 1500);
    });
  }

  return {
    element: controls,
    dock,
    pad,
    /** Reveal the controls — hosts keep them hidden until the arcade boots. */
    show: () => { dock.hidden = false; },
    isCompact: () => compact,
    setPad,
    ui: { carts, playpause, restart, pace, rendering, stats, hint, pad, setPad },
    destroy: () => {
      observer.disconnect();
      coarse.removeEventListener('change', remeasure);
      controls.remove();
    },
  };
}
