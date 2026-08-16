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
  return `${type}<w:headerReference w:type="default" r:id="rId${10 + index * 2}"/>`
    + `<w:footerReference w:type="default" r:id="rId${11 + index * 2}"/>`
    + `<w:pgSz w:w="${width}" w:h="${height}"${width > height ? ' w:orient="landscape"' : ""}/>`
    + `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" `
    + `w:header="720" w:footer="720" w:gutter="0"/>${columns}${numbering}`;
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

/** Minimal deterministic document whose only declared face is supplied by the #442 test matrix. */
export function generateFontProbeDocx(
  family = "Docxodus Canvas Mono",
  text = "AZ AZA ZAZ",
  {
    pageWidth = 12240,
    pageHeight = 15840,
    margin = 1440,
    paragraphCount = 1,
  } = {},
) {
  const paragraphs = Array.from({ length: paragraphCount }, () => `
  <w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr>
    <w:t>${text}</w:t></w:r></w:p>`).join("");
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
    <w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>${paragraphs}
  <w:sectPr><w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}"/></w:sectPr>
</w:body></w:document>`),
    },
  ]);
}
