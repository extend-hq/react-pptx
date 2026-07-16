import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import PptxGenJS from 'pptxgenjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'tests/fixtures/pretext-layout-cases.pptx');
const scratch = mkdtempSync(join(tmpdir(), 'pretext-layout-fixture-'));
const base = join(scratch, 'base.pptx');
const unpacked = join(scratch, 'unpacked');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'react-pptx';
pptx.subject = 'PowerPoint text-layout regression fixture';
pptx.title = 'Pretext layout stress deck';
pptx.company = 'Extend';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Arial',
  bodyFontFace: 'Arial',
  lang: 'en-US',
};

const black = '111111';
const gray = '5B6470';
const blue = '3987FF';
const pale = 'F5F7F9';

function title(slide, text, kicker = 'PRETEXT LAYOUT FIXTURE') {
  slide.background = { color: 'FFFFFF' };
  slide.addText(kicker, {
    x: 0.42,
    y: 0.14,
    w: 5.4,
    h: 0.22,
    fontFace: 'Arial',
    fontSize: 8,
    bold: true,
    color: gray,
    margin: 0,
  });
  slide.addText(text, {
    x: 0.42,
    y: 0.46,
    w: 12.35,
    h: 0.55,
    fontFace: 'Arial',
    fontSize: 25,
    color: black,
    margin: 0,
    breakLine: false,
  });
}

function card(slide, x, y, w, h, name, fill = pale) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    name,
    rectRadius: 0.06,
    fill: { color: fill },
    line: { color: 'C2C8D0', width: 0.8 },
  });
}

function page(slide, number) {
  slide.addText(String(number), {
    x: 12.65,
    y: 7.12,
    w: 0.2,
    h: 0.18,
    fontSize: 8,
    color: gray,
    margin: 0,
    align: 'right',
  });
}

const paragraph =
  'PowerPoint lays out a paragraph by combining the selected font, kerning, line-breaking rules, and the available width. When a line is justified, the remaining width is distributed across eligible spaces. Small metric differences accumulate across a sentence, changing both the final word position and the next wrap point.';

{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  slide.addText('TEXT LAYOUT DIAGNOSTIC', {
    x: 0.42,
    y: 0.45,
    w: 4,
    h: 0.25,
    fontSize: 12,
    bold: true,
    color: gray,
    margin: 0,
  });
  slide.addText('Pretext layout\nstress deck', {
    x: 0.42,
    y: 2.7,
    w: 6.8,
    h: 1.55,
    fontSize: 44,
    color: black,
    margin: 0,
    breakLine: false,
  });
  slide.addText(
    'Eight focused cases for justification, run boundaries, tabs, hanging indents, scripts, columns, and autofit.',
    { x: 0.42, y: 5.15, w: 7.9, h: 0.6, fontSize: 17, color: black, margin: 0 },
  );
}

{
  const slide = pptx.addSlide();
  title(slide, 'Justification exposes the word-spacing gap');
  card(slide, 0.42, 1.82, 5.95, 4.5, 'JustifiedCard');
  card(slide, 6.92, 1.82, 5.95, 4.5, 'DistributedCard');
  slide.addText('JUSTIFIED — algn=just', {
    x: 0.67,
    y: 2.1,
    w: 3,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  slide.addText(paragraph, {
    x: 0.67,
    y: 2.62,
    w: 5.45,
    h: 2.8,
    name: 'JustifiedText',
    fontSize: 16,
    color: black,
    margin: 0,
    breakLine: false,
  });
  slide.addText('DISTRIBUTED — algn=dist', {
    x: 7.17,
    y: 2.1,
    w: 3.3,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  slide.addText(paragraph, {
    x: 7.17,
    y: 2.62,
    w: 5.45,
    h: 2.8,
    name: 'DistributedText',
    fontSize: 16,
    color: black,
    margin: 0,
    breakLine: false,
  });
  page(slide, 2);
}

{
  const slide = pptx.addSlide();
  title(slide, 'Formatting runs should not move the words');
  slide.addText('ONE UNIFORM RUN', {
    x: 0.42,
    y: 1.8,
    w: 3,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: gray,
    margin: 0,
  });
  slide.addText(paragraph, {
    x: 0.42,
    y: 2.32,
    w: 5.65,
    h: 3.2,
    name: 'UniformRuns',
    fontSize: 17,
    color: black,
    margin: 0,
    breakLine: false,
  });
  slide.addText('EQUIVALENT TEXT SPLIT ACROSS RUNS', {
    x: 6.83,
    y: 1.8,
    w: 4,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: gray,
    margin: 0,
  });
  slide.addText(
    [
      { text: 'PowerPoint ', options: { bold: true } },
      { text: 'should preserve ', options: {} },
      { text: 'identical wrap points ', options: { italic: true } },
      {
        text: 'even when an otherwise uniform sentence is split into many formatting runs. ',
        options: {},
      },
      { text: 'Boundaries ', options: { bold: true } },
      {
        text: 'introduced for emphasis, hyperlinks, or language metadata must not create extra spacing.',
        options: {},
      },
    ],
    {
      x: 6.83,
      y: 2.32,
      w: 5.65,
      h: 3.2,
      name: 'SplitRuns',
      fontSize: 17,
      color: black,
      margin: 0,
      breakLine: false,
    },
  );
  page(slide, 3);
}

{
  const slide = pptx.addSlide();
  title(slide, 'Tabs and hanging indents require paragraph geometry');
  card(slide, 0.42, 1.78, 7.9, 4.6, 'TabsCard');
  slide.addText('EXPLICIT TAB STOPS — 1.35 in / 3.00 in / 4.70 in', {
    x: 0.72,
    y: 2.08,
    w: 5,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  slide.addText(
    'Item\tOwner\tStatus\nText measurement\tLayout engine\tIn progress\nScript font selection\tFont resolver\tPlanned\nPowerPoint oracle\tOracle export\tReady',
    {
      x: 0.72,
      y: 2.62,
      w: 7.25,
      h: 2.3,
      name: 'ExplicitTabs',
      fontSize: 16,
      color: black,
      margin: 0,
      breakLine: false,
    },
  );
  slide.addText('HANGING INDENT', {
    x: 8.78,
    y: 2.08,
    w: 2.2,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  slide.addText(
    'First: preserve the exact marker and text origin across wrapped lines.\nSecond: retain the same hanging geometry after a hard break.',
    {
      x: 8.95,
      y: 2.62,
      w: 3.7,
      h: 2.4,
      name: 'HangingIndent',
      fontSize: 16,
      color: black,
      margin: 0,
      breakLine: false,
    },
  );
  page(slide, 4);
}

{
  const slide = pptx.addSlide();
  title(slide, 'Script-aware wrapping changes the break opportunities');
  const boxes = [
    [
      0.42,
      'LatinText',
      'LATIN — keep words intact',
      'Internationalization requires predictable line breaking, even when a narrow text box contains exceptionallylongtechnicalidentifiers.',
    ],
    [
      4.7,
      'JapaneseText',
      'JAPANESE — East Asian breaks',
      'プレゼンテーションの文字組みでは、句読点や英数字を含む長い文章でも、自然な位置で折り返す必要があります。',
    ],
    [
      8.98,
      'ArabicText',
      'ARABIC — RTL + complex script',
      'اختبار الشرائح في النص العربي يتطلب اتجاه الكتابة الصحيح والخط المناسب والمحافظة على علامات الكلمات.',
    ],
  ];
  for (const [x, name, label, text] of boxes) {
    card(
      slide,
      Number(x),
      1.82,
      3.86,
      4.5,
      `${name}Card`,
      name === 'JapaneseText' ? 'EAF5FC' : pale,
    );
    slide.addText(String(label), {
      x: Number(x) + 0.22,
      y: 2.18,
      w: 3.4,
      h: 0.28,
      fontSize: 10,
      bold: true,
      color: name === 'JapaneseText' ? blue : gray,
      margin: 0,
    });
    slide.addText(String(text), {
      x: Number(x) + 0.22,
      y: 2.82,
      w: 3.42,
      h: 2.7,
      name: String(name),
      fontSize: 15,
      color: black,
      margin: 0,
      breakLine: false,
      rtlMode: name === 'ArabicText',
    });
  }
  page(slide, 5);
}

{
  const slide = pptx.addSlide();
  title(slide, 'Columns and autofit depend on measured line results');
  slide.addText('TWO-COLUMN TEXT BODY', {
    x: 0.42,
    y: 1.82,
    w: 3,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  card(slide, 0.42, 2.12, 5.95, 4.2, 'ColumnsCard');
  slide.addText(`${paragraph} ${paragraph}`, {
    x: 0.68,
    y: 2.42,
    w: 5.45,
    h: 3.55,
    name: 'TwoColumns',
    fontSize: 14,
    color: black,
    margin: 0,
    breakLine: false,
  });
  slide.addText('NORMAL AUTOFIT — SAME BOX, LONG COPY', {
    x: 6.92,
    y: 1.82,
    w: 4.2,
    h: 0.25,
    fontSize: 10,
    bold: true,
    color: blue,
    margin: 0,
  });
  card(slide, 6.92, 2.12, 5.95, 4.2, 'AutofitCard', 'EAF5FC');
  slide.addText(`${paragraph} ${paragraph}`, {
    x: 7.18,
    y: 2.42,
    w: 5.45,
    h: 3.55,
    name: 'NormalAutofit',
    fontSize: 16,
    color: black,
    margin: 0,
    breakLine: false,
  });
  page(slide, 6);
}

{
  const slide = pptx.addSlide();
  title(
    slide,
    'A synthetic image exclusion shows the Pretext opportunity',
    'PROPOSED VIEWER LAYOUT — NOT NATIVE PPTX WRAP METADATA',
  );
  card(slide, 2.1, 1.64, 9.15, 4.88, 'SyntheticWrapCard');
  slide.addText(
    'Pretext can route successive lines through different horizontal intervals while an image temporarily removes part of the available line width.',
    { x: 2.45, y: 2.02, w: 8.4, h: 0.7, fontSize: 14, color: black, margin: 0 },
  );
  slide.addText(
    'Lines beside the image use a narrower interval without inserting manual source breaks.',
    { x: 2.45, y: 3.02, w: 2.65, h: 1.4, fontSize: 14, color: black, margin: 0 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 5.3,
    y: 2.98,
    w: 2.72,
    h: 1.9,
    name: 'SyntheticImage',
    fill: { color: '1B1831' },
    line: { color: '1B1831' },
  });
  slide.addText('IMAGE\nEXCLUSION', {
    x: 5.55,
    y: 3.58,
    w: 2.2,
    h: 0.5,
    fontSize: 15,
    bold: true,
    color: 'FFFFFF',
    align: 'center',
    margin: 0,
  });
  slide.addText(
    'The same logical line can continue in a second interval when the layout policy allows fragmented rows.',
    { x: 8.35, y: 3.02, w: 2.55, h: 1.4, fontSize: 14, color: black, margin: 0 },
  );
  slide.addText('The PPTX uses separate boxes only to illustrate the target geometry.', {
    x: 2.45,
    y: 5.35,
    w: 8.4,
    h: 0.45,
    fontSize: 13,
    color: black,
    margin: 0,
  });
  page(slide, 7);
}

{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  slide.addText('PASS CRITERIA', {
    x: 0.42,
    y: 0.45,
    w: 3,
    h: 0.25,
    fontSize: 12,
    bold: true,
    color: gray,
    margin: 0,
  });
  slide.addText('The words finish\nwhere PowerPoint puts them', {
    x: 0.42,
    y: 2.75,
    w: 9.2,
    h: 1.5,
    fontSize: 40,
    color: black,
    margin: 0,
    breakLine: false,
  });
  slide.addText(
    'Same line count  •  Same wrap words  •  Same final line width  •  Same column break',
    { x: 0.42, y: 5.28, w: 11.4, h: 0.38, fontSize: 17, color: black, margin: 0 },
  );
}

function mergeAttributes(raw, attributes) {
  let next = raw;
  for (const [name, value] of Object.entries(attributes)) {
    next = next.replace(new RegExp(`\\s${name}="[^"]*"`, 'g'), '');
    next += ` ${name}="${value}"`;
  }
  return next;
}

function mapShape(slideNumber, markerText, shapeName, mapper, occurrence = 0) {
  const path = join(unpacked, `ppt/slides/slide${slideNumber}.xml`);
  const xml = readFileSync(path, 'utf8');
  let markerIndex = -1;
  for (let index = 0; index <= occurrence; index += 1) {
    markerIndex = xml.indexOf(markerText, markerIndex + 1);
  }
  if (markerIndex < 0) throw new Error(`Missing ${markerText} on slide ${slideNumber}`);
  const start = xml.lastIndexOf('<p:sp>', markerIndex);
  const end = xml.indexOf('</p:sp>', markerIndex) + '</p:sp>'.length;
  if (start < 0 || end < '</p:sp>'.length) throw new Error(`Malformed ${shapeName}`);
  const mapped = mapper(xml.slice(start, end)).replace(
    /(<p:cNvPr[^>]*\sname=")[^"]*/,
    `$1${shapeName}`,
  );
  writeFileSync(path, `${xml.slice(0, start)}${mapped}${xml.slice(end)}`);
}

function paragraphProperties(fragment, attributes, children = '') {
  let found = false;
  let next = fragment.replace(/<a:pPr([^>]*)\/>/g, (_match, raw) => {
    found = true;
    return `<a:pPr${mergeAttributes(raw, attributes)}>${children}</a:pPr>`;
  });
  if (found) return next;
  next = next.replace(/<a:pPr([^>]*)>/g, (_match, raw) => {
    found = true;
    return `<a:pPr${mergeAttributes(raw, attributes)}>${children}`;
  });
  return found
    ? next
    : next.replace(/<a:p>/g, `<a:p><a:pPr${mergeAttributes('', attributes)}>${children}</a:pPr>`);
}

function runProperties(fragment, attributes, fonts = '') {
  return fragment.replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_match, raw, children) => {
    const withoutFonts = children.replace(/<a:(?:latin|ea|cs|sym)[^>]*\/>/g, '');
    return `<a:rPr${mergeAttributes(raw, attributes)}>${fonts}${withoutFonts}</a:rPr>`;
  });
}

function bodyProperties(fragment, attributes, child = '') {
  let next = fragment.replace(
    /<a:bodyPr([^>]*)\/>/,
    (_match, raw) => `<a:bodyPr${mergeAttributes(raw, attributes)}>${child}</a:bodyPr>`,
  );
  if (next === fragment) {
    next = fragment.replace(/<a:bodyPr([^>]*)>([\s\S]*?)<\/a:bodyPr>/, (_match, raw, children) => {
      const cleaned = children.replace(/<a:(?:noAutofit|normAutofit|spAutoFit)[^>]*\/>/g, '');
      return `<a:bodyPr${mergeAttributes(raw, attributes)}>${child}${cleaned}</a:bodyPr>`;
    });
  }
  return next;
}

await pptx.writeFile({ fileName: base });
execFileSync('unzip', ['-q', base, '-d', unpacked]);

mapShape(2, 'PowerPoint lays out', 'JustifiedText', (shape) =>
  paragraphProperties(shape, { algn: 'just' }),
);
mapShape(
  2,
  'PowerPoint lays out',
  'DistributedText',
  (shape) => paragraphProperties(shape, { algn: 'dist' }),
  1,
);
mapShape(3, 'PowerPoint lays out', 'UniformRuns', (shape) =>
  runProperties(shape, { kern: '1200', spc: '75', lang: 'en-US' }),
);
mapShape(3, 'should preserve', 'SplitRuns', (shape) =>
  runProperties(shape, { kern: '1200', spc: '75', lang: 'en-US' }),
);
mapShape(4, 'Item', 'ExplicitTabs', (shape) =>
  paragraphProperties(
    shape,
    { defTabSz: '914400' },
    '<a:tabLst><a:tab pos="1234440" algn="l"/><a:tab pos="2743200" algn="l"/><a:tab pos="4297680" algn="l"/></a:tabLst>',
  ),
);
mapShape(4, 'First:', 'HangingIndent', (shape) =>
  paragraphProperties(shape, { marL: '400050', indent: '-228600' }),
);
mapShape(5, 'Internationalization', 'LatinText', (shape) =>
  paragraphProperties(shape, { latinLnBrk: '0' }),
);
mapShape(5, 'プレゼンテーション', 'JapaneseText', (shape) => {
  const paragraph = paragraphProperties(shape, { eaLnBrk: '1', hangingPunct: '1' });
  return runProperties(
    paragraph,
    { lang: 'ja-JP', altLang: 'en-US', kern: '1200' },
    '<a:latin typeface="Arial"/><a:ea typeface="Yu Gothic"/>',
  );
});
mapShape(5, 'اختبار الشرائح', 'ArabicText', (shape) => {
  const paragraph = paragraphProperties(shape, { rtl: '1' });
  return runProperties(
    paragraph,
    { lang: 'ar-SA', kern: '1200' },
    '<a:latin typeface="Arial"/><a:cs typeface="Noto Naskh Arabic"/><a:rtl/>',
  );
});
mapShape(6, 'PowerPoint lays out', 'TwoColumns', (shape) =>
  bodyProperties(shape, { numCol: '2', spcCol: '182880' }),
);
mapShape(
  6,
  'PowerPoint lays out',
  'NormalAutofit',
  (shape) => bodyProperties(shape, {}, '<a:normAutofit fontScale="88000" lnSpcReduction="12000"/>'),
  2,
);

rmSync(output, { force: true });
execFileSync('zip', ['-q', '-r', output, '.'], { cwd: unpacked });
rmSync(scratch, { force: true, recursive: true });
console.log(output);
