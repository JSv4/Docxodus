import { drawingDocx } from './docx-drawing-shape-fixture.js';

export type ShapeWrap = 'none' | 'topAndBottom' | 'square' | 'tight' | 'through';
export interface FloatingShapeOptions {
  grouped?: boolean;
  wrap?: ShapeWrap;
  behind?: boolean;
  order?: number;
  mixed?: boolean;
  geometry?: string;
  horizontal?: 'column' | 'page';
  vertical?: 'paragraph' | 'page';
  x?: number;
  y?: number;
  distance?: number;
  bottomDistance?: number;
  topDistance?: number;
  relativeWidth?: number;
  relativeHeight?: number;
  heading?: boolean;
  outline?: boolean;
  story?: 'header' | 'footer';
  height?: number;
  before?: string[];
  after?: string;
  layers?: { order: number; fill: string; behind?: boolean; height?: number }[];
}
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS='http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const emu=(pt:number)=>Math.round(pt*12700);
const parts = [
  {id:'rStyles',type:R+'/styles',name:'styles.xml',contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    body:`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="atLeast"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`},
  {id:'rSettings',type:R+'/settings',name:'settings.xml',contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',body:`<w:settings xmlns:w="${W}"/>`},
];

export function vmlShapeDocx(withExtent = true): Uint8Array {
  return drawingDocx('<w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" '
    + (withExtent ? 'style="width:144pt;height:72pt" ' : '')
    + 'fillcolor="#19647E"/></w:pict></w:r>',false,parts);
}

export function floatingShapeDocx(options: FloatingShapeOptions = {}): Uint8Array {
  const source = (id:number, order:number, fill:string, behind:boolean, height = emu(options.height ?? 72)) => {
    const shape=`<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="${height}"/></a:xfrm>`
      + `<a:prstGeom prst="${options.geometry??'rect'}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
      + (options.outline?'<a:ln w="19050"><a:solidFill><a:srgbClr val="FF9900"/></a:solidFill></a:ln>':'')
      + '</wps:spPr><wps:bodyPr/></wps:wsp>';
    const content = options.grouped ? '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
      + `<a:ext cx="1828800" cy="${height}"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="${height}"/></a:xfrm></wpg:grpSpPr>`
      + shape+'</wpg:wgp>' : shape;
    const kind=options.wrap??'topAndBottom';
    const polygon='<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="10800"/>'
      + '<wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
    const wrap=kind==='none'?'<wp:wrapNone/>':kind==='topAndBottom'?'<wp:wrapTopAndBottom/>':kind==='square'
      ? '<wp:wrapSquare wrapText="bothSides"/>' : `<wp:wrap${kind==='tight'?'Tight':'Through'} wrapText="bothSides">${polygon}</wp:wrap${kind==='tight'?'Tight':'Through'}>`;
    const relative=(options.relativeWidth===undefined?'':`<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>${options.relativeWidth}</wp14:pctWidth></wp14:sizeRelH>`)
      +(options.relativeHeight===undefined?'':`<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>${options.relativeHeight}</wp14:pctHeight></wp14:sizeRelV>`);
    return `<w:r><w:drawing xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}"`
      +(relative?' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"':'')+'>'
      + `<wp:anchor distT="${emu(options.topDistance??0)}" distR="${emu(options.distance??0)}" distB="${emu(options.bottomDistance??0)}" distL="0" simplePos="0" relativeHeight="${order}" behindDoc="${behind?'1':'0'}" locked="0" layoutInCell="1" allowOverlap="1">`
      + '<wp:simplePos x="0" y="0"/>'
      + `<wp:positionH relativeFrom="${options.horizontal??'column'}"><wp:posOffset>${emu(options.x??0)}</wp:posOffset></wp:positionH>`
      + `<wp:positionV relativeFrom="${options.vertical??'paragraph'}"><wp:posOffset>${emu(options.y??0)}</wp:posOffset></wp:positionV>`
      + `<wp:extent cx="1828800" cy="${height}"/>${wrap}<wp:docPr id="${id}" name="Harbor panel ${id}"/>`
      + `<a:graphic><a:graphicData uri="${options.grouped?WPG:WPS}">${content}</a:graphicData></a:graphic>${relative}</wp:anchor></w:drawing></w:r>`;
  };
  const layers=options.layers??[{order:options.order??10,fill:'19647E',behind:options.behind??false}];
  const drawings=layers.map((l,i)=>source(7+i,l.order,l.fill,l.behind??false,l.height === undefined ? undefined : emu(l.height))).join('');
  const headingStyle='<w:style w:type="paragraph" w:styleId="DrawingHeading"><w:name w:val="Drawing heading"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';
  const fixtureParts=options.heading?parts.map(p=>p.name==='styles.xml'?{...p,body:p.body.replace('</w:styles>',headingStyle+'</w:styles>')}:p):[...parts];
  if(options.story) fixtureParts.push({id:'rDrawingStory',type:R+'/'+options.story,name:options.story+'1.xml',
    contentType:`application/vnd.openxmlformats-officedocument.wordprocessingml.${options.story}+xml`,
    body:`<w:${options.story==='header'?'hdr':'ftr'} xmlns:w="${W}"><w:p>${drawings}</w:p></w:${options.story==='header'?'hdr':'ftr'}>`});
  return drawingDocx((options.heading?'<w:pPr><w:pStyle w:val="DrawingHeading"/></w:pPr>':'')
    +(options.story?'<w:r><w:t>StoryBody</w:t></w:r>':drawings)+(options.mixed?'<w:r><w:t>AnchorWord</w:t></w:r>':''),
    false,fixtureParts,options.after,options.before,options.story?`<w:${options.story}Reference w:type="default" r:id="rDrawingStory"/>`:'');
}
