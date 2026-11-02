import { insetClipBounds } from './line-metrics.js';

const TABLE = 'data-docx-exact-table-borders';
const CELL = 'data-docx-exact-cell-borders';
const EDGES = 'data-docx-table-border-edges';
const PREPARED = 'data-docx-table-borders-prepared';
const PT = 4/3;
type Edge = ['top' | 'bottom' | 'vertical', number, number, number, string];
const edgesOf = (row: HTMLTableRowElement): Edge[] => JSON.parse(row.getAttribute(EDGES) ?? '[]');

/** A continuation must reserve its actual outer edges before the paginator measures its fit. */
export function refreshTableBorderLayout(table: HTMLTableElement): void {
  if (!table.hasAttribute(PREPARED) || !table.rows.length) return;
  const rows = Array.from(table.rows);
  if (table.hasAttribute('data-docx-grid-align-end')) {
    const end = Number(table.getAttribute(TABLE))+table.querySelectorAll(':scope > colgroup > col').length;
    const half = Math.max(0,...rows.flatMap(row => edgesOf(row))
      .filter(edge => edge[0] === 'vertical' && edge[1] === end).map(edge => edge[3]))*PT/2;
    table.style.setProperty(`margin-${table.getAttribute('data-docx-grid-margin-side')}`,`${half}px`);
  }
  const firstBody = rows.findIndex(row => row.dataset.docxRepeatedHeaderRow !== 'true');
  const join = firstBody > 0 ? rows[firstBody].getAttribute('data-docx-table-header-edges') : null;
  if (join) {
    const edges = JSON.parse(join) as Edge[];
    rows[firstBody].setAttribute(EDGES,JSON.stringify([
      ...edgesOf(rows[firstBody]).filter(e => e[0] !== 'top'),...edges,
    ]));
    rows[firstBody-1].setAttribute(EDGES,JSON.stringify([
      ...edgesOf(rows[firstBody-1]).filter(e => e[0] !== 'bottom'),
      ...edges.map(e => ['bottom',...e.slice(1)]),
    ]));
  }
  const half = (row: HTMLTableRowElement, side: string) =>
    Math.max(0,...edgesOf(row).filter(e => e[0] === side).map(e => e[3]))*PT/2;
  for (const [index,row] of rows.entries()) {
    const height = Number(row.getAttribute('data-docx-source-row-height'))*PT;
    const rule = row.getAttribute('data-docx-row-height-rule');
    for (const cell of Array.from(row.cells)) {
      if (!cell.hasAttribute(CELL)) continue;
      const last = cell.rowSpan === 0 ? rows.length-1 : Math.min(rows.length-1,index+cell.rowSpan-1);
      const top = (Number.parseFloat(cell.style.getPropertyValue('--docx-cell-base-padding-top')) || 0)+half(row,'top');
      const bottom = (Number.parseFloat(cell.style.getPropertyValue('--docx-cell-base-padding-bottom')) || 0)+half(rows[last],'bottom');
      cell.style.paddingTop = `${top}px`; cell.style.paddingBottom = `${bottom}px`;
      const clip = cell.querySelector<HTMLElement>(':scope > [data-docx-exact-row-cell]');
      if (height > 0 && rule === 'exact' && clip) clip.style.height = `${Math.max(0,height-top-bottom)}px`;
    }
    if (height > 0 && (rule === 'exact' || rule === 'atLeast'))
      row.style.height = `${height+(rule === 'exact' ? 0 : half(row,'top')+half(row,'bottom'))}px`;
  }
  for (const [side, row] of [['top', table.rows[0]], ['bottom', table.rows[table.rows.length-1]]] as const) {
    const base = Number.parseFloat(table.style.getPropertyValue(`--docx-table-base-margin-${side}`)) || 0;
    table.style.setProperty(`margin-${side}`, `${base+half(row,side)}px`);
  }
}

/** Keep native borders in ordinary HTML; replace their quantized boxes before export measures. */
export function prepareTableBorders(root: HTMLElement): void {
  const view = root.ownerDocument.defaultView;
  if (!view) return;
  for (const table of Array.from(root.querySelectorAll<HTMLTableElement>(`table[${TABLE}]`))) {
    if (table.hasAttribute(PREPARED)) continue;
    const style = view.getComputedStyle(table);
    const points = (name: string) => (Number.parseFloat(style.getPropertyValue(name)) || 0)*PT;
    const marginTop = Number.parseFloat(style.marginTop) || 0, marginBottom = Number.parseFloat(style.marginBottom) || 0;
    const side = table.getAttribute('data-docx-grid-margin-side');
    if (side) table.style.setProperty(`margin-${side}`, `${points('--docx-table-grid-margin')}px`);
    table.style.setProperty('--docx-table-base-margin-top', `${marginTop}px`);
    table.style.setProperty('--docx-table-base-margin-bottom', `${marginBottom}px`);
    for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>(`td[${CELL}]`))) {
      if (cell.closest('table') !== table) continue;
      const cellStyle = view.getComputedStyle(cell);
      const padding: Record<string, number> = {};
      for (const edge of ['top','right','bottom','left']) {
        const half = (Number.parseFloat(cellStyle.getPropertyValue(`--docx-cell-border-${edge}`)) || 0)*PT/2;
        const margin = Number.parseFloat(cellStyle.getPropertyValue(`padding-${edge}`)) || 0;
        if (edge === 'top' || edge === 'bottom') cell.style.setProperty(`--docx-cell-base-padding-${edge}`,`${margin}px`);
        padding[edge] = edge === 'left' || edge === 'right' ? Math.max(margin,half) : margin+half;
      }
      cell.style.border = 'none';
      for (const edge of Object.keys(padding)) cell.style.setProperty(`padding-${edge}`, `${padding[edge]}px`);
      cell.style.cssText = cell.style.cssText;
    }
    table.style.cssText = table.style.cssText;
    table.setAttribute(PREPARED,'true');
    refreshTableBorderLayout(table);
  }
}

/** The padding-box coordinate system, including every ancestor rotation, reflection and scale. */
function paddingMatrix(element: HTMLElement, view: Window & typeof globalThis): DOMMatrix {
  const saved = element.style.cssText;
  const probes: HTMLElement[] = [];
  try {
    if (view.getComputedStyle(element).position === 'static') {
      element.style.position = 'relative';
      element.style.inset = 'auto';
    }
    const points = [[0,0],[100,0],[0,100]].map(([x,y]) => {
      const probe = element.ownerDocument.createElement('span');
      probe.style.cssText = `all:initial;display:block;position:absolute;left:${x}px;top:${y}px;`
        + 'width:0;height:0;min-width:0;min-height:0;margin:0;padding:0;border:0;visibility:hidden;pointer-events:none';
      element.append(probe); probes.push(probe);
      return probe.getBoundingClientRect();
    });
    const [origin,x,y] = points;
    return new view.DOMMatrix([(x.left-origin.left)/100,(x.top-origin.top)/100,
      (y.left-origin.left)/100,(y.top-origin.top)/100,origin.left,origin.top]);
  } finally {
    probes.forEach(probe => probe.remove());
    element.style.cssText = saved;
  }
}

const matrixText = (m: DOMMatrix) => `matrix(${[m.a,m.b,m.c,m.d,m.e,m.f].join(' ')})`;
const SVG_NS = 'http://www.w3.org/2000/svg';

/** Carry the original ancestor clips into the overlay, each in its own transformed space. */
function clippedGroup(table: HTMLTableElement, page: HTMLElement, svg: SVGSVGElement,
  toOverlay: DOMMatrix, view: Window & typeof globalThis): SVGGElement {
  const doc = table.ownerDocument;
  let parent: SVGElement = svg;
  const appendClip = (shapes: SVGElement[]) => {
    const clip = doc.createElementNS(SVG_NS,'clipPath');
    let serial = svg.childElementCount;
    while (doc.getElementById(`docx-table-clip-${serial}`)) serial++;
    clip.id = `docx-table-clip-${serial}`;
    clip.setAttribute('clipPathUnits','userSpaceOnUse');
    clip.append(...shapes); svg.append(clip);
    const group = doc.createElementNS(SVG_NS,'g');
    group.setAttribute('clip-path',`url(#${clip.id})`);
    parent.append(group); parent = group;
  };
  for (let ancestor: Element | null = table; ancestor && ancestor !== page; ancestor = ancestor.parentElement) {
    if (ancestor instanceof view.SVGGraphicsElement) {
      // Exact-row text preservation uses an SVG path with zero-area spikes. Copy that actual
      // clip, including its pixel-snapped visible rectangle, instead of its bounding rectangle.
      const id = /^url\(["']?#([^"')]+)["']?\)$/.exec(ancestor.getAttribute('clip-path') ?? '')?.[1];
      const source = id ? doc.getElementById(id) : null, ctm = ancestor.getScreenCTM();
      if (source instanceof view.SVGClipPathElement && ctm) {
        const matrix = matrixText(toOverlay.multiply(new view.DOMMatrix([ctm.a,ctm.b,ctm.c,ctm.d,ctm.e,ctm.f])));
        appendClip(Array.from(source.children).map(child => {
          const clone = child.cloneNode(true) as SVGElement;
          clone.removeAttribute('id');
          clone.setAttribute('transform',`${matrix} ${clone.getAttribute('transform') ?? ''}`);
          return clone;
        }));
      }
      continue;
    }
    if (!(ancestor instanceof view.HTMLElement)) continue;
    const style = view.getComputedStyle(ancestor);
    const clipsX = /^(hidden|clip|scroll|auto)$/.test(style.overflowX);
    const clipsY = /^(hidden|clip|scroll|auto)$/.test(style.overflowY);
    if (!clipsX && !clipsY && style.clipPath === 'none') continue;
    const px = (name: string) => Number.parseFloat(style.getPropertyValue(name)) || 0;
    const bl = px('border-left-width'), bt = px('border-top-width');
    const br = px('border-right-width'), bb = px('border-bottom-width');
    const w = px('width') + (style.boxSizing === 'border-box' ? -bl-br : px('padding-left')+px('padding-right'));
    const h = px('height') + (style.boxSizing === 'border-box' ? -bt-bb : px('padding-top')+px('padding-bottom'));
    const reach = 100000;
    const margin = px('overflow-clip-margin');
    let left = clipsX ? -margin : -reach, right = clipsX ? w+margin : reach;
    let top = clipsY ? -margin : -reach, bottom = clipsY ? h+margin : reach;
    const inset = insetClipBounds(style.clipPath,new view.DOMRect(-bl,-bt,w+bl+br,h+bt+bb),1);
    if (inset) {
      left = Math.max(left,inset.left); right = Math.min(right,inset.right);
      top = Math.max(top,inset.top); bottom = Math.min(bottom,inset.bottom);
    }
    const rect = doc.createElementNS(SVG_NS,'rect');
    const matrix = toOverlay.multiply(paddingMatrix(ancestor,view));
    for (const [name,value] of Object.entries({x:left,y:top,width:Math.max(0,right-left),height:Math.max(0,bottom-top)}))
      rect.setAttribute(name,String(value));
    rect.setAttribute('transform',matrixText(matrix));
    appendClip([rect]);
  }
  const group = doc.createElementNS(SVG_NS,'g');
  parent.append(group);
  return group;
}

/** Draw final fragment geometry while retaining the table's original stacking ancestry and order. */
export function paintTableBorders(container: HTMLElement): void {
  const doc = container.ownerDocument, view = doc.defaultView;
  if (!view) return;
  for (const table of Array.from(container.querySelectorAll<HTMLTableElement>(`table[${PREPARED}]`))) {
    const page = table.closest<HTMLElement>('.page-box');
    if (!page) continue;
    const matrix = paddingMatrix(table,view), inverse = matrix.inverse();
    if (![inverse.a,inverse.b,inverse.c,inverse.d,inverse.e,inverse.f].every(Number.isFinite)) continue;
    const columns = Array.from(table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col'));
    const boundaries = [0];
    for (const column of columns) boundaries.push(boundaries[boundaries.length-1]+Number.parseFloat(view.getComputedStyle(column).width));
    const prefix = Number(table.getAttribute(TABLE));
    const rtl = view.getComputedStyle(table).direction === 'rtl';
    const xAt = (index: number) => rtl ? boundaries[boundaries.length-1]-boundaries[index-prefix] : boundaries[index-prefix];
    const svg = doc.createElementNS(SVG_NS,'svg');
    svg.setAttribute('aria-hidden','true');
    svg.setAttribute('data-docx-table-border-paint','true');
    svg.setAttribute('width',String(page.offsetWidth));
    svg.setAttribute('height',String(page.offsetHeight));
    svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;pointer-events:none';
    // A page-level overlay would escape a textbox's stacking context and paint over later
    // foreground drawings. Adjacent placement preserves that order without changing any
    // ancestor's positioning or the containing block of its authored absolute descendants.
    table.before(svg);
    const screen = svg.getScreenCTM();
    if (!screen) { svg.remove(); continue; }
    // Chromium exposes the legacy SVGMatrix here; normalize it before mixing with DOMMatrix.
    const toOverlay = new view.DOMMatrix([screen.a,screen.b,screen.c,screen.d,screen.e,screen.f]).inverse();
    const group = clippedGroup(table,page,svg,toOverlay,view);
    group.setAttribute('transform',matrixText(toOverlay.multiply(matrix)));
    const rows = Array.from(table.rows), painted = new Set<string>();
    const allEdges = rows.map(edgesOf);
    for (const [index,row] of rows.entries()) {
      const bounds = row.getBoundingClientRect();
      const center = inverse.transformPoint({x:(bounds.left+bounds.right)/2,y:(bounds.top+bounds.bottom)/2});
      const height = Number.parseFloat(view.getComputedStyle(row).height);
      const top = center.y-height/2, bottom = center.y+height/2;
      for (const [side,start,end,width,color] of allEdges[index]) {
        const stroke = width*PT;
        let x: number, y: number, w: number, h: number;
        if (side === 'vertical') {
          x = xAt(start)-stroke/2; y = top; w = stroke; h = height;
        } else {
          // Extend through the vertical halves at a join, closing outer corners and T junctions.
          const adjacent = allEdges[index+(side === 'top' ? -1 : 1)] ?? [];
          const half = (column: number) => Math.max(0,...[...allEdges[index],...adjacent]
            .filter(e => e[0] === 'vertical' && e[1] === column).map(e => e[3]))*PT/2;
          const a = xAt(start), b = xAt(end), startHalf = half(start), endHalf = half(end);
          x = Math.min(a,b)-(rtl ? endHalf : startHalf);
          w = Math.abs(b-a)+startHalf+endHalf;
          y = (side === 'top' ? top : bottom)-stroke/2; h = stroke;
        }
        const box = [x,y,w,h];
        if (!box.every(Number.isFinite) || !(w > 0 && h > 0)) continue;
        const fill = color === 'currentColor' ? view.getComputedStyle(row).color : color;
        // Adjacent rows share a source boundary even when independent matrix inversion gives
        // their coordinates a few floating-point ulps of difference. Paint that boundary once.
        const key = `${side === 'vertical' ? `row:${index}` : `boundary:${side === 'top' ? index : index+1}`}|${start}|${end}|${width}|${fill}`;
        if (painted.has(key)) continue;
        painted.add(key);
        const rect = doc.createElementNS(SVG_NS,'rect');
        ['x','y','width','height'].forEach((name,i) => rect.setAttribute(name,String(box[i])));
        rect.setAttribute('fill',fill); group.append(rect);
      }
    }
  }
}
