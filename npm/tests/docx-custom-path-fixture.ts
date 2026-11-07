import { inlineDrawingDocx } from './docx-drawing-shape-fixture.js';

const pt = (x: number | string, y: number) => `<a:pt x="${x}" y="${y}"/>`;
const move = (x: number | string, y: number) => `<a:moveTo>${pt(x, y)}</a:moveTo>`;
const line = (x: number, y: number) => `<a:lnTo>${pt(x, y)}</a:lnTo>`;
const polygon = move(100, 450) + line(500, 50) + line(900, 450) + '<a:close/>';
const path = (commands: string, paint = 'stroke="0"', dimensions = 'w="1000" h="500"') =>
  `<a:path ${dimensions} ${paint}>${commands}</a:path>`;

export const CUSTOM_PATHS = {
  polygon: path(polygon),
  quadratic: path(move(100, 400) + `<a:quadBezTo>${pt(500, -200)}${pt(900, 400)}</a:quadBezTo>`
    + line(900, 450) + line(100, 450) + '<a:close/>'),
  cubic: path(move(100, 400) + `<a:cubicBezTo>${pt(200, 0)}${pt(800, 0)}${pt(900, 400)}</a:cubicBezTo>`
    + line(100, 400) + '<a:close/>'),
  open: path(move(100, 100) + line(900, 400), 'fill="none"'),
  hole: path(move(100, 50) + line(900, 50) + line(900, 450) + line(100, 450) + '<a:close/>'
    + move(300, 150) + line(300, 350) + line(700, 350) + line(700, 150) + '<a:close/>'),
  sameWinding: path(move(100, 50) + line(900, 50) + line(900, 450) + line(100, 450) + '<a:close/>'
    + move(300, 150) + line(700, 150) + line(700, 350) + line(300, 350) + '<a:close/>'),
  multiple: path(polygon) + path(move(200, 500) + line(1800, 500), 'fill="none"', 'w="2000" h="1000"'),
  invisible: path(polygon, 'fill="none" stroke="false"'),
  arc: path(polygon) + path(move(900, 250) + '<a:arcTo wR="400" hR="200" stAng="0" swAng="10800000"/>'),
  guide: path(move('wd2', 50) + line(900, 450) + line(100, 450) + '<a:close/>'),
  dimensions: path(polygon, 'stroke="0"', ''),
  fill: path(polygon, 'fill="darken" stroke="0"'),
};

export function customPathDocx(kind: keyof typeof CUSTOM_PATHS, grouped = false,
  options: { text?: boolean; effect?: boolean; transform?: string; opacity?: 'alpha' | 'alphaMod' } = {}): Uint8Array {
  const shape = '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
    + `<a:xfrm ${options.transform ?? ''}><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>`
    + '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="100" t="100" r="1000" b="1000"/>'
    + `<a:pathLst>${CUSTOM_PATHS[kind]}</a:pathLst></a:custGeom>`
    + '<a:solidFill><a:srgbClr val="19647E">'
    + (options.opacity ? `<a:${options.opacity} val="${options.opacity === 'alpha' ? 40000 : 0}"/>` : '')
    + '</a:srgbClr></a:solidFill>'
    + '<a:ln w="25400"><a:solidFill><a:srgbClr val="883311"/></a:solidFill></a:ln>'
    + (options.effect ? '<a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>' : '')
    + '</wps:spPr>'
    + (options.text ? '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Custom caption</w:t></w:r></w:p></w:txbxContent></wps:txbx>' : '')
    + '<wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>';
  const group = '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
    + '<a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>'
    + `</a:xfrm></wpg:grpSpPr>${shape}</wpg:wgp>`;
  return inlineDrawingDocx('http://schemas.microsoft.com/office/word/2010/wordprocessing' + (grouped ? 'Group' : 'Shape'),
    grouped ? group : shape);
}
