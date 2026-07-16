import type { PresentationDocument } from '@extend-ai/react-pptx-model';
import { afterEach, describe, expect, it } from 'vitest';
import { NormalizedPresentationViewer } from './normalized-viewer';

const documentModel: PresentationDocument = {
  format: 'pptx',
  size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
  slides: [
    {
      id: 'slide-1',
      index: 0,
      nodes: [
        {
          id: 'layout-cases',
          type: 'shape',
          transform: { x: 0, y: 0, width: 7_620_000, height: 4_000_000 },
          geometry: { preset: 'rect' },
          columnCount: 2,
          columnSpacing: 182_880,
          paragraphs: [
            {
              alignment: 'distributed',
              defaultTabSizeEmu: 914_400,
              tabStops: [
                { positionEmu: 1_234_440, alignment: 'left' },
                { positionEmu: 2_743_200, alignment: 'left' },
                { positionEmu: 4_297_680, alignment: 'left' },
              ],
              runs: [
                {
                  text: 'Item\tOwner\tStatus\nText measurement\tLayout engine\tIn progress',
                  fontFamily: 'Aptos',
                },
              ],
            },
            {
              bullet: { kind: 'character', value: '•' },
              marginLeftEmu: 400_050,
              indentEmu: -228_600,
              runs: [{ text: 'Hanging bullet' }],
            },
            {
              eastAsianLineBreak: false,
              hangingPunctuation: true,
              runs: [
                {
                  text: '日本語',
                  fontFamily: 'Aptos',
                  eastAsianFontFamily: 'Yu Gothic',
                  complexScriptFontFamily: 'Arial',
                  language: 'ja-JP',
                },
              ],
            },
            {
              latinLineBreak: true,
              runs: [
                {
                  text: 'العربية',
                  fontFamily: 'Aptos',
                  complexScriptFontFamily: 'Noto Naskh Arabic',
                  language: 'ar-SA',
                  rightToLeft: true,
                  fontSizePt: 18,
                  kerningThresholdPt: 12,
                },
                {
                  text: ' small',
                  fontSizePt: 10,
                  kerningThresholdPt: 12,
                },
              ],
            },
          ],
        },
      ],
    },
  ],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [],
};

describe('DrawingML text layout semantics', () => {
  afterEach(() => document.body.replaceChildren());

  it('renders explicit tabs, script controls, kerning, hanging indents, and columns', async () => {
    const container = document.createElement('div');
    document.body.append(container);
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const content = container.querySelector<HTMLElement>('[data-rpv-text-content]')!;
    expect(content.style.columnCount).toBe('2');
    expect(content.style.columnGap).toBe('19.2px');
    expect(content.style.width).toBe('100%');
    expect(content.style.height).toBe('100%');

    const paragraphs = [...container.querySelectorAll<HTMLElement>('[data-rpv-text-paragraph]')];
    expect(paragraphs[0]!.style.textAlign).toBe('justify');
    expect(paragraphs[0]!.style.textAlignLast).toBe('justify');
    expect(paragraphs[0]!.style.getPropertyValue('text-justify')).toBe('inter-character');
    expect(paragraphs[0]!.style.tabSize).toBe('96px');
    const rows = paragraphs[0]!.querySelectorAll<HTMLElement>('[data-rpv-tab-row]');
    expect(rows).toHaveLength(2);
    expect(rows[0]!.style.gridTemplateColumns).toBe('129.6px 158.4px 163.2px minmax(0, 1fr)');
    expect(rows[0]!.querySelectorAll('[data-rpv-tab-cell]')).toHaveLength(3);
    expect(rows[1]!.querySelectorAll('[data-rpv-tab-cell]')).toHaveLength(3);

    const marker = paragraphs[1]!.querySelector<HTMLElement>('[data-rpv-bullet]')!;
    expect(paragraphs[1]!.style.marginLeft).toBe('42px');
    expect(paragraphs[1]!.style.textIndent).toBe('-24px');
    expect(marker.style.display).toBe('inline-block');
    expect(marker.style.width).toBe('24px');

    const japanese = [...paragraphs[2]!.querySelectorAll<HTMLElement>('span')].find(
      (span) => span.textContent === '日本語',
    )!;
    expect(japanese.style.fontFamily).toBe('Yu Gothic, Aptos, Arial');
    expect(japanese.style.wordBreak).toBe('keep-all');
    expect(paragraphs[2]!.style.getPropertyValue('hanging-punctuation')).toBe(
      'first allow-end last',
    );

    const arabic = [...paragraphs[3]!.querySelectorAll<HTMLElement>('span')].find(
      (span) => span.textContent === 'العربية',
    )!;
    const small = [...paragraphs[3]!.querySelectorAll<HTMLElement>('span')].find(
      (span) => span.textContent === ' small',
    )!;
    expect(arabic.style.fontFamily).toBe('Noto Naskh Arabic, Aptos');
    expect(arabic.dir).toBe('rtl');
    expect(arabic.style.unicodeBidi).toBe('embed');
    expect(arabic.style.fontKerning).toBe('normal');
    expect(arabic.style.overflowWrap).toBe('anywhere');
    expect(small.style.fontKerning).toBe('none');

    viewer.destroy();
  });
});
