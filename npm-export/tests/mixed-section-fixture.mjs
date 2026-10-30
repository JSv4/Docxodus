const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function xml(value) {
  return Buffer.from(value, "utf8");
}

function storedZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const checksum = crc32(entry.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(entry.data.length, 18);
    local.writeUInt32LE(entry.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    localParts.push(local, name, entry.data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(entry.data.length, 20);
    central.writeUInt32LE(entry.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, name);
    offset += local.length + name.length + entry.data.length;
  }
  const directory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...localParts, directory, end]));
}

function paragraph(text, extraRuns = "") {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`
    + `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>${extraRuns}</w:p>`;
}

/** A new, minimal document whose complete PDF must survive owned-browser shutdown. */
export function generateBrowserCleanupDocx() {
  return storedZip([
    { name: "[Content_Types].xml", data: xml(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>') },
    { name: "_rels/.rels", data: xml(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`
      + '</Relationships>') },
    { name: "word/document.xml", data: xml(
      `<w:document xmlns:w="${W_NS}"><w:body>`
      + paragraph("CEDAR WORKSHOP INSPECTION")
      + paragraph("Seven tools were returned to cabinet 24.")
      + paragraph("Inspection complete. Signed: Rowan Vale.")
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
      + '</w:sectPr></w:body></w:document>') },
  ]);
}

function pageBreak() {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>`
    + `<w:r><w:br w:type="page"/></w:r></w:p>`;
}

function sectionProperties(index, width, height, options = {}) {
  const type = options.type ? `<w:type w:val="${options.type}"/>` : "";
  const columns = options.columns === 2
    ? `<w:cols w:num="2" w:space="720"/>`
    : `<w:cols w:space="720"/>`;
  const numbering = options.start === undefined
    ? ""
    : `<w:pgNumType w:start="${options.start}" w:fmt="decimal"/>`;
  // CT_SectPr is ordered: story references, type, page size/margins, numbering, then columns.
  // Keeping generated fixtures schema-valid prevents a tolerant converter from masking an invalid
  // section sequence that Word or another conforming consumer may repair differently.
  return `<w:headerReference w:type="default" r:id="rId${10 + index * 2}"/>`
    + `<w:footerReference w:type="default" r:id="rId${11 + index * 2}"/>`
    + type
    + `<w:pgSz w:w="${width}" w:h="${height}"${width > height ? ' w:orient="landscape"' : ""}/>`
    + `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" `
    + `w:header="720" w:footer="720" w:gutter="0"/>${numbering}${columns}`;
}

function sectionParagraph(text, extraRuns, index, width, height, options) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`
    + `<w:sectPr>${sectionProperties(index, width, height, options)}</w:sectPr></w:pPr>`
    + `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>${extraRuns}</w:p>`;
}

function field(instruction, cached) {
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`
    + `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>`
    + `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`
    + `<w:r><w:t>${cached}</w:t></w:r>`
    + `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

function headerPart(index) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<w:hdr xmlns:w="${W_NS}">${paragraph(`HEADER-S${index}`)}</w:hdr>`;
}

function footerPart(index) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<w:ftr xmlns:w="${W_NS}"><w:p><w:r><w:t xml:space="preserve">FOOTER-S${index} PAGE </w:t></w:r>`
    + `${field("PAGE", "1")}</w:p></w:ftr>`;
}

/**
 * Readable six-page #440 fixture:
 * Letter portrait x2 -> Letter landscape -> A4 portrait shared with a continuous two-column
 * section -> A4 spill -> Letter portrait. The page break in section 3 makes the continuous spill
 * deterministic without relying on host font line wrapping.
 */
export function generateMixedSectionDocx() {
  const contentOverrides = Array.from({ length: 5 }, (_, index) =>
    `<Override PartName="/word/header${index + 1}.xml" `
      + `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`
      + `<Override PartName="/word/footer${index + 1}.xml" `
      + `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
    .join("\n  ");
  const storyRelationships = Array.from({ length: 5 }, (_, index) =>
    `<Relationship Id="rId${10 + index * 2}" Type="${R_NS}/header" Target="header${index + 1}.xml"/>`
      + `<Relationship Id="rId${11 + index * 2}" Type="${R_NS}/footer" Target="footer${index + 1}.xml"/>`)
    .join("\n  ");
  const entries = [
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  ${contentOverrides}
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/_rels/document.xml.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
  ${storyRelationships}
</Relationships>`),
    },
    {
      name: "word/styles.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`),
    },
    {
      name: "word/footnotes.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> LANDSCAPE FOOTNOTE TOKEN</w:t></w:r></w:p></w:footnote>
</w:footnotes>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${paragraph("BODY-S0-P1 LETTER PORTRAIT")}
  ${pageBreak()}
  ${sectionParagraph("BODY-S0-P2 EXPLICIT PAGE BREAK", "", 0, 12240, 15840, { start: 1 })}
  ${sectionParagraph("BODY-S1 LANDSCAPE", `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`, 1, 15840, 12240, { start: 10 })}
  ${sectionParagraph("BODY-S2 A4 PORTRAIT BEFORE CONTINUOUS", "", 2, 11906, 16838)}
  ${paragraph("BODY-S3 TWO COLUMN SHARED PAGE")}
  ${pageBreak()}
  ${sectionParagraph("BODY-S3 TWO COLUMN SPILL PAGE", "", 3, 11906, 16838, { type: "continuous", columns: 2 })}
  ${paragraph("BODY-S4 LETTER PORTRAIT FINAL")}
  <w:sectPr>${sectionProperties(4, 12240, 15840)}</w:sectPr>
</w:body></w:document>`),
    },
  ];

  for (let index = 0; index < 5; index++) {
    entries.push({ name: `word/header${index + 1}.xml`, data: xml(headerPart(index)) });
    entries.push({ name: `word/footer${index + 1}.xml`, data: xml(footerPart(index)) });
  }
  return storedZip(entries);
}

/** One real-converter footnote paragraph long enough to span several PDF pages. */
export function generateLongFootnoteDocx(wordCount = 600) {
  const noteText = Array.from(
    { length: wordCount },
    (_, index) => `footnote-1-1-${index + 1}`,
  ).join(" ");
  return storedZip([
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/_rels/document.xml.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
</Relationships>`),
    },
    {
      name: "word/styles.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
</w:styles>`),
    },
    {
      name: "word/footnotes.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> ${noteText}</w:t></w:r>
  </w:p></w:footnote>
</w:footnotes>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:t>Body cites long footnote</w:t></w:r>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
  </w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

/** One paragraph carrying a native deletion and insertion, for review-profile assertions. */
export function generateTrackedRevisionDocx() {
  return storedZip([
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/_rels/document.xml.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: "word/styles.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>
    <w:del w:id="1" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>
    <w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

export function exactRowCell(body, properties = "") {
  return `<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/>${properties}</w:tcPr>${body}</w:tc>`;
}

export function exactRowParagraph(text, properties = "", spacing = 'w:line="240" w:lineRule="auto"') {
  return `<w:p><w:pPr>${properties}<w:spacing w:before="0" w:after="0" ${spacing}/></w:pPr>`
    + `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/**
 * A table whose first row has an exact height (`w:hRule="exact"`) too short for its first cell, so
 * the cell's second paragraph is clipped as Word clips it. Paragraphs before and after the table
 * and an ordinary second row pin the order of the PDF's text. The options move the table to other
 * page geometries and cell shapes:
 * - `pageWidth`, `pageHeight`, `margins` (twips) and `leadingParagraphs` (filler paragraphs before the table);
 * - `vAlign` for the clipped cell, `rowHeight` (twips) for the exact row, `spacing` for every paragraph;
 * - `firstRow`/`secondRow` replace a row's `<w:tr>` outright (`exactRowCell`, `exactRowParagraph` build them);
 * - `extraParts`/`extraRelationships`/`extraContentTypes` add media the replaced rows refer to.
 */
export function generateExactRowOverflowDocx(options = {}) {
  const {
    pageWidth = 12240,
    pageHeight = 15840,
    margins = { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    leadingParagraphs = 0,
    vAlign,
    rowHeight = 360,
    spacing,
    extraParts = [],
    extraRelationships = "",
    extraContentTypes = "",
  } = options;
  const p = (text) => exactRowParagraph(text, "", spacing);
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((side) => `<w:${side} w:val="single" w:sz="6" w:color="444444"/>`).join("");
  const alignment = vAlign ? `<w:vAlign w:val="${vAlign}"/>` : "";
  const firstRow = options.firstRow ?? `<w:tr><w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="exact"/></w:trPr>`
    + exactRowCell(p("Birch exact record.") + p("Overflow Birch continuation line."), alignment)
    + exactRowCell(p("Neighbour cell."))
    + `</w:tr>`;
  const secondRow = options.secondRow ?? `<w:tr>${exactRowCell(p("Following row text."))}`
    + `${exactRowCell(p("Second column."))}</w:tr>`;
  const filler = Array.from({ length: leadingParagraphs }, (_, index) =>
    p(`Filler paragraph ${index + 1} keeps the table further down the document.`)).join("");
  return storedZip([
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>${extraContentTypes}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/_rels/document.xml.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>${extraRelationships}
</Relationships>`),
    },
    {
      name: "word/styles.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
  ${filler}${p("Opening ledger paragraph.")}
  <w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>
    <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>
    ${firstRow}
    ${secondRow}
  </w:tbl>
  ${p("Closing ledger paragraph.")}
  <w:sectPr><w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"/><w:pgMar w:top="${margins.top}" w:right="${margins.right}" w:bottom="${margins.bottom}" w:left="${margins.left}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`),
    },
    ...extraParts,
  ]);
}
