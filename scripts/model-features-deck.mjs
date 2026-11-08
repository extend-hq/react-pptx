// Builds a small, hand-written PPTX that exercises the presentation-model
// fields PowerPoint writes outside plain shapes and text: slide transitions,
// animation timing, Office Math, text fields, merged table cells, click
// actions, and accessibility metadata. The parser contract tests parse the
// bytes directly and the fixture generator also writes them to disk.

const PML = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const DECORATIVE = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const MODEL_FEATURES_DECK_NAME = 'model-features.pptx';
export const SLIDE_NUMBER_FIELD_ID = '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}';

const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

function emptyGroup() {
  return (
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  );
}

function run(text) {
  return `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${text}</a:t></a:r>`;
}

function textShape(id, name, x, y, cx, cy, paragraphs, nonVisual = '') {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"${nonVisual}<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
  );
}

function tableCell(text, attributes = '') {
  return (
    `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${text ? run(text) : ''}</a:p></a:txBody><a:tcPr/></a:tc>`
  );
}

function slideOne() {
  const equation =
    '<a:p>' +
    run('Energy: ') +
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="a14"><a14:m xmlns:a14="${A14}" xmlns:m="${MATH}">` +
    '<m:oMathPara><m:oMath><m:sSup><m:sSupPr/><m:e><m:r><m:t>E=mc</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>' +
    '</a14:m></mc:Choice><mc:Fallback>' +
    run('E=mc²') +
    '</mc:Fallback></mc:AlternateContent></a:p>';
  const slideNumber =
    `<a:p><a:pPr algn="r"/><a:fld id="${SLIDE_NUMBER_FIELD_ID}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p>`;
  const nextButton =
    ' descr="Advances to the next slide" title="Next"><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></p:cNvPr>';
  const summaryLink = '><a:hlinkClick r:id="rId2" action="ppaction://hlinksldjump"/></p:cNvPr>';
  const table =
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Merged table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="914400" y="3657600"/><a:ext cx="5486400" cy="1371600"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/>' +
    '<a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>' +
    `<a:tr h="457200">${tableCell('Quarter', ' gridSpan="2"')}${tableCell('', ' hMerge="1"')}</a:tr>` +
    `<a:tr h="457200">${tableCell('North', ' rowSpan="2"')}${tableCell('Q1')}</a:tr>` +
    `<a:tr h="457200">${tableCell('', ' vMerge="1"')}${tableCell('Q2')}</a:tr>` +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  const transition =
    '<p:transition spd="slow" advClick="1" advTm="3000"><p:fade thruBlk="1"/></p:transition>';
  const timing =
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
    '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' +
    '<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    '<p:par><p:cTn id="5" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    '<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>' +
    '<p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>' +
    '<p:to><p:strVal val="visible"/></p:to></p:set>' +
    '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>' +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
    '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="2" grpId="0"/></p:bldLst></p:timing>';

  return (
    xmlHeader +
    `<p:sld xmlns:a="${DML}" xmlns:r="${REL}" xmlns:p="${PML}"><p:cSld name="Model features"><p:spTree>` +
    emptyGroup() +
    textShape(2, 'Equation', 914400, 914400, 6096000, 914400, equation, '/>') +
    textShape(3, 'Slide number', 914400, 2286000, 2743200, 457200, slideNumber, '/>') +
    textShape(4, 'Next button', 4572000, 2286000, 2743200, 457200, `<a:p>${run('Next')}</a:p>`, nextButton) +
    textShape(5, 'Summary link', 7772400, 2286000, 2743200, 457200, `<a:p>${run('Jump to summary')}</a:p>`, summaryLink) +
    table +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    transition +
    timing +
    '</p:sld>'
  );
}

function slideTwo() {
  const decorative =
    `><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="${DECORATIVE}" val="1"/></a:ext></a:extLst></p:cNvPr>`;
  return (
    xmlHeader +
    `<p:sld xmlns:a="${DML}" xmlns:r="${REL}" xmlns:p="${PML}"><p:cSld name="Summary"><p:spTree>` +
    emptyGroup() +
    textShape(2, 'Summary title', 914400, 914400, 9144000, 914400, `<a:p>${run('Summary')}</a:p>`, '/>') +
    textShape(3, 'Hidden note', 914400, 2286000, 4572000, 457200, `<a:p>${run('Speaker-only reminder')}</a:p>`, ' hidden="1"/>') +
    textShape(4, 'Decorative bar', 914400, 3200400, 9144000, 228600, '<a:p/>', decorative) +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '<p:transition spd="med"><p:push dir="l"/></p:transition>' +
    '</p:sld>'
  );
}

function theme() {
  const color = (name, value) => `<a:${name}><a:srgbClr val="${value}"/></a:${name}>`;
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (w) => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  return (
    xmlHeader +
    `<a:theme xmlns:a="${DML}" name="Fixture"><a:themeElements><a:clrScheme name="Fixture">` +
    color('dk1', '000000') + color('lt1', 'FFFFFF') + color('dk2', '44546A') + color('lt2', 'E7E6E6') +
    color('accent1', '4472C4') + color('accent2', 'ED7D31') + color('accent3', 'A5A5A5') +
    color('accent4', 'FFC000') + color('accent5', '5B9BD5') + color('accent6', '70AD47') +
    color('hlink', '0563C1') + color('folHlink', '954F72') +
    '</a:clrScheme><a:fontScheme name="Fixture">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme><a:fmtScheme name="Fixture">' +
    `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
    '</a:fmtScheme></a:themeElements></a:theme>'
  );
}

function slideMaster() {
  return (
    xmlHeader +
    `<p:sldMaster xmlns:a="${DML}" xmlns:r="${REL}" xmlns:p="${PML}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    `<p:spTree>${emptyGroup()}</p:spTree></p:cSld>` +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'
  );
}

function slideLayout() {
  return (
    xmlHeader +
    `<p:sldLayout xmlns:a="${DML}" xmlns:r="${REL}" xmlns:p="${PML}" type="blank" preserve="1"><p:cSld name="Blank">` +
    `<p:spTree>${emptyGroup()}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

function presentation() {
  return (
    xmlHeader +
    `<p:presentation xmlns:a="${DML}" xmlns:r="${REL}" xmlns:p="${PML}" saveSubsetFonts="1">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>' +
    '<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>' +
    '</p:presentation>'
  );
}

function relationships(entries) {
  return (
    xmlHeader +
    `<Relationships xmlns="${PKG_REL}">` +
    entries
      .map(
        ([id, type, target, mode]) =>
          `<Relationship Id="${id}" Type="${REL_TYPE}/${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`,
      )
      .join('') +
    '</Relationships>'
  );
}

function contentTypes() {
  const override = (part, type) =>
    `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${type}+xml"/>`;
  return (
    xmlHeader +
    `<Types xmlns="${CT}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    override('ppt/presentation.xml', 'presentation.main') +
    override('ppt/slideMasters/slideMaster1.xml', 'slideMaster') +
    override('ppt/slideLayouts/slideLayout1.xml', 'slideLayout') +
    override('ppt/slides/slide1.xml', 'slide') +
    override('ppt/slides/slide2.xml', 'slide') +
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
    '</Types>'
  );
}

/** The package parts as `[path, xml]` pairs, in a deterministic order. */
export function modelFeaturesParts() {
  return [
    ['[Content_Types].xml', contentTypes()],
    ['_rels/.rels', relationships([['rId1', 'officeDocument', 'ppt/presentation.xml']])],
    ['ppt/presentation.xml', presentation()],
    [
      'ppt/_rels/presentation.xml.rels',
      relationships([
        ['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'],
        ['rId2', 'slide', 'slides/slide1.xml'],
        ['rId3', 'slide', 'slides/slide2.xml'],
        ['rId4', 'theme', 'theme/theme1.xml'],
      ]),
    ],
    ['ppt/slideMasters/slideMaster1.xml', slideMaster()],
    [
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      relationships([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
        ['rId2', 'theme', '../theme/theme1.xml'],
      ]),
    ],
    ['ppt/slideLayouts/slideLayout1.xml', slideLayout()],
    [
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      relationships([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    ],
    ['ppt/theme/theme1.xml', theme()],
    ['ppt/slides/slide1.xml', slideOne()],
    [
      'ppt/slides/_rels/slide1.xml.rels',
      relationships([
        ['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
        ['rId2', 'slide', 'slide2.xml'],
      ]),
    ],
    ['ppt/slides/slide2.xml', slideTwo()],
    [
      'ppt/slides/_rels/slide2.xml.rels',
      relationships([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
    ],
  ];
}

const crcTable = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Writes the parts into a stored (uncompressed) ZIP archive with fixed
 * timestamps, so the same input always yields byte-identical output.
 */
export function buildModelFeaturesDeck() {
  const encoder = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;
  // 1980-01-01 00:00:00 in MS-DOS date/time form.
  const dosTime = 0;
  const dosDate = (1 << 5) | 1;
  for (const [path, xml] of modelFeaturesParts()) {
    const name = encoder.encode(path);
    const data = encoder.encode(xml);
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0, true);
    local.setUint16(8, 0, true);
    local.setUint16(10, dosTime, true);
    local.setUint16(12, dosDate, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, name.length, true);
    local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);

    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(4, 20, true);
    entry.setUint16(6, 20, true);
    entry.setUint16(8, 0, true);
    entry.setUint16(10, 0, true);
    entry.setUint16(12, dosTime, true);
    entry.setUint16(14, dosDate, true);
    entry.setUint32(16, crc, true);
    entry.setUint32(20, data.length, true);
    entry.setUint32(24, data.length, true);
    entry.setUint16(28, name.length, true);
    entry.setUint16(30, 0, true);
    entry.setUint16(32, 0, true);
    entry.setUint16(34, 0, true);
    entry.setUint16(36, 0, true);
    entry.setUint32(38, 0, true);
    entry.setUint32(42, offset, true);
    central.push(new Uint8Array(entry.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(4, 0, true);
  end.setUint16(6, 0, true);
  end.setUint16(8, central.length / 2, true);
  end.setUint16(10, central.length / 2, true);
  end.setUint32(12, centralSize, true);
  end.setUint32(16, offset, true);
  end.setUint16(20, 0, true);
  const all = [...chunks, ...central, new Uint8Array(end.buffer)];
  const output = new Uint8Array(all.reduce((sum, chunk) => sum + chunk.length, 0));
  let position = 0;
  for (const chunk of all) {
    output.set(chunk, position);
    position += chunk.length;
  }
  return output;
}
