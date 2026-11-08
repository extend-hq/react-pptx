import type {
  ImageNode,
  MediaNode,
  PresentationDocument,
  ShapeNode,
  SlideNode,
  TableNode,
} from '@extend-ai/react-pptx-model';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { NormalizedPresentationViewer } from './normalized-viewer';

const MATHML = 'http://www.w3.org/1998/Math/MathML';
const SVG = 'http://www.w3.org/2000/svg';
const transform = { x: 0, y: 0, width: 1_000_000, height: 500_000 };

function presentation(nodes: SlideNode[], extraSlides: Array<Partial<PresentationDocument['slides'][number]>> = []): PresentationDocument {
  return {
    format: 'pptx',
    size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
    slides: [
      { id: 'slide-1', index: 0, nodes, sourcePart: 'ppt/slides/slide1.xml' },
      ...extraSlides.map((slide, offset) => ({
        id: `slide-${offset + 2}`,
        index: offset + 1,
        nodes: [],
        sourcePart: `ppt/slides/slide${offset + 2}.xml`,
        ...slide,
      })),
    ],
    masters: [],
    layouts: [],
    themes: [],
    assets: {},
    warnings: [],
  };
}

function shape(id: string, text: string, extra: Partial<ShapeNode> = {}): ShapeNode {
  return {
    id,
    type: 'shape',
    transform,
    geometry: { preset: 'rect' },
    paragraphs: [{ runs: [{ text }] }],
    ...extra,
  };
}

async function render(document: PresentationDocument) {
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const viewer = new NormalizedPresentationViewer(container, document);
  await viewer.renderSlide(0);
  return { container, viewer };
}

describe('normalized viewer model features', () => {
  beforeEach(() => {
    vi.stubGlobal('URL', Object.assign(URL, {
      createObjectURL: vi.fn(() => 'blob:mock-asset'),
      revokeObjectURL: vi.fn(),
    }));
  });

  afterEach(() => {
    window.document.body.replaceChildren();
    delete (HTMLImageElement.prototype as { decode?: unknown }).decode;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('renders Office Math runs as sanitized MathML and keeps them upright', async () => {
    const mathMl =
      `<math xmlns="${MATHML}" display="inline"><msup><mi>E</mi><mn>2</mn></msup><mo onclick="alert(1)">=</mo></math>`;
    const { container, viewer } = await render(
      presentation([
        shape('math', 'E=mc2', {
          paragraphs: [
            {
              runs: [
                { text: 'Energy: ' },
                { text: 'E=mc2', italic: true, math: { text: 'E=mc2', display: true, sourceXml: '', mathMl } },
              ],
            },
          ],
        }),
        shape('bad-math', 'fallback', {
          paragraphs: [
            {
              runs: [
                {
                  text: 'x',
                  math: {
                    text: 'x',
                    display: false,
                    sourceXml: '',
                    mathMl: `<math xmlns="${MATHML}"><semantics><mi>x</mi></semantics></math>`,
                  },
                },
              ],
            },
          ],
        }),
      ]),
    );

    const mathSpan = container.querySelector<HTMLElement>('[data-rpv-node-id="math"] [data-rpv-math]')!;
    const math = mathSpan.querySelector('math')!;
    expect(math.namespaceURI).toBe(MATHML);
    expect(math.querySelector('msup > mi')?.textContent).toBe('E');
    expect(math.querySelector('mo')?.hasAttribute('onclick')).toBe(false);
    expect(mathSpan.style.fontStyle).toBe('normal');
    expect(container.querySelector('[data-rpv-node-id="math"]')?.textContent).toContain('Energy: ');

    // MathML the viewer does not allow-list falls back to the cached text.
    const fallback = container.querySelector('[data-rpv-node-id="bad-math"]')!;
    expect(fallback.querySelector('math')).toBeNull();
    expect(fallback.textContent).toContain('x');
    viewer.destroy();
  });

  it('skips merged table cells and keeps the spans on their origin cells', async () => {
    const table: TableNode = {
      id: 'table',
      type: 'table',
      transform,
      columnWidths: [500_000, 500_000],
      rowHeights: [100_000, 100_000, 100_000],
      rows: [
        [
          { colSpan: 2, paragraphs: [{ runs: [{ text: 'Quarter' }] }] },
          { merge: { horizontal: true, vertical: false, origin: { row: 0, column: 0 } }, paragraphs: [{ runs: [{ text: 'hidden' }] }] },
        ],
        [
          { rowSpan: 2, paragraphs: [{ runs: [{ text: 'North' }] }] },
          { paragraphs: [{ runs: [{ text: 'Q1' }] }] },
        ],
        [
          { merge: { horizontal: false, vertical: true, origin: { row: 1, column: 0 } }, paragraphs: [] },
          { paragraphs: [{ runs: [{ text: 'Q2' }] }] },
        ],
      ],
    };
    const { container, viewer } = await render(presentation([table]));

    const cells = [...container.querySelectorAll('[data-rpv-node-id="table"] td')];
    expect(cells.map((cell) => cell.textContent?.trim())).toEqual(['Quarter', 'North', 'Q1', 'Q2']);
    expect(cells[0]).toHaveProperty('colSpan', 2);
    expect(cells[1]).toHaveProperty('rowSpan', 2);
    expect(container.textContent).not.toContain('hidden');
    expect(viewer.searchText('hidden')).toEqual([]);
    viewer.destroy();
  });

  it('turns presentation jump and slide-link actions into in-viewer navigation', async () => {
    const { container, viewer } = await render(
      presentation(
        [
          shape('next', 'Next', {
            clickAction: { relationshipId: '', action: 'ppaction://hlinkshowjump?jump=nextslide' },
            accessibilityTitle: 'Next',
            description: 'Advances to the next slide',
          }),
          shape('last', 'Last', { clickAction: { action: 'ppaction://hlinkshowjump?jump=lastslide' } }),
          shape('summary', 'Summary', {
            hyperlink: 'ppt/slides/slide3.xml',
            clickAction: { external: false, relationshipId: 'rId2', target: 'ppt/slides/slide3.xml', action: 'ppaction://hlinksldjump' },
          }),
          shape('web', 'Docs', { hyperlink: 'https://example.com/docs' }),
        ],
        [{}, {}],
      ),
    );
    const goToSlide = vi.spyOn(viewer, 'goToSlide').mockResolvedValue();

    const next = container.querySelector<HTMLElement>('[data-rpv-node-id="next"]')!;
    expect(next.tagName).toBe('DIV');
    expect(next.getAttribute('role')).toBe('link');
    expect(next.tabIndex).toBe(0);
    expect(next.title).toBe('Next');
    next.click();
    expect(goToSlide).toHaveBeenLastCalledWith(1);

    container.querySelector<HTMLElement>('[data-rpv-node-id="last"]')!.click();
    expect(goToSlide).toHaveBeenLastCalledWith(2);

    const summary = container.querySelector<HTMLElement>('[data-rpv-node-id="summary"]')!;
    expect(summary.tagName).toBe('DIV');
    expect(summary.getAttribute('href')).toBeNull();
    summary.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(goToSlide).toHaveBeenLastCalledWith(2);

    const web = container.querySelector<HTMLAnchorElement>('[data-rpv-node-id="web"]')!;
    expect(web.tagName).toBe('A');
    expect(web.href).toBe('https://example.com/docs');
    expect(goToSlide).toHaveBeenCalledTimes(3);
    viewer.destroy();
  });

  it('applies hidden and decorative metadata to the node element', async () => {
    const { container, viewer } = await render(
      presentation([
        shape('hidden', 'Speaker note', { hidden: true }),
        shape('decor', '', { decorative: true, paragraphs: [] }),
        shape('plain', 'Visible'),
      ]),
    );
    expect(container.querySelector<HTMLElement>('[data-rpv-node-id="hidden"]')!.hidden).toBe(true);
    expect(container.querySelector('[data-rpv-node-id="decor"]')!.getAttribute('aria-hidden')).toBe('true');
    expect(container.querySelector('[data-rpv-node-id="plain"]')!.getAttribute('aria-hidden')).toBeNull();
    viewer.destroy();
  });

  it('shows a media poster with a play control that reveals the trimmed video', async () => {
    const play = vi.fn(() => Promise.resolve());
    vi.spyOn(HTMLMediaElement.prototype, 'play').mockImplementation(play);
    const poster: ImageNode = {
      id: 'poster',
      type: 'image',
      transform,
      assetId: '',
      pictureSvg: `<svg xmlns="${SVG}" viewBox="0 0 10 10"><rect width="10" height="10" fill="#888"/></svg>`,
    };
    const media: MediaNode = {
      id: 'clip',
      type: 'media',
      transform,
      mediaType: 'video',
      altText: 'Launch video',
      sourceUrl: 'https://media.example.com/launch.mp4',
      poster,
      playback: { trimStartMs: 1500, trimEndMs: 500 },
    };
    const { container, viewer } = await render(presentation([media]));

    const node = container.querySelector<HTMLElement>('[data-rpv-node-id="clip"]')!;
    const video = node.querySelector('video')!;
    const button = node.querySelector('button')!;
    expect(node.querySelector('[data-rpv-node-id="poster"] img')?.getAttribute('src')).toMatch(/^data:image\/svg\+xml/);
    expect(video.hidden).toBe(true);
    expect(video.getAttribute('src')).toBe('https://media.example.com/launch.mp4');
    expect(video.preload).toBe('none');
    expect(button.getAttribute('aria-label')).toBe('Play Launch video');

    button.click();
    expect(play).toHaveBeenCalledTimes(1);
    expect(video.hidden).toBe(false);
    expect(node.querySelector('button')).toBeNull();
    expect(node.querySelector<HTMLElement>('[data-rpv-node-id="poster"]')!.hidden).toBe(true);

    // Trim points clamp playback to the saved window once metadata is known.
    Object.defineProperty(video, 'duration', { value: 10, configurable: true });
    video.dispatchEvent(new Event('loadedmetadata'));
    expect(video.currentTime).toBe(1.5);
    video.currentTime = 9.8;
    video.dispatchEvent(new Event('seeking'));
    expect(video.currentTime).toBe(9.5);
    viewer.destroy();
  });

  it('offers embedded OLE objects as downloads named after their asset', async () => {
    const document = presentation([
      {
        id: 'sheet',
        type: 'media',
        transform,
        mediaType: 'ole',
        assetId: 'ole-1',
        embeddedObject: { programId: 'Excel.Sheet.12', name: 'Budget', linked: false, showAsIcon: true },
      } satisfies MediaNode,
    ]);
    document.assets['ole-1'] = {
      id: 'ole-1',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      byteLength: 3,
      fileName: 'budget.xlsx',
      data: new Uint8Array([1, 2, 3]),
    };
    const { container, viewer } = await render(document);

    const node = container.querySelector<HTMLElement>('[data-rpv-node-id="sheet"]')!;
    expect(node.dataset.rpvEmbeddedObject).toBe('Excel.Sheet.12');
    const download = node.querySelector<HTMLAnchorElement>('a')!;
    expect(download.download).toBe('budget.xlsx');
    expect(download.getAttribute('href')).toBe('blob:mock-asset');
    expect(node.querySelector('video, audio')).toBeNull();
    viewer.destroy();
  });

  it('paints the fallback picture when the primary artwork is missing', async () => {
    // jsdom does not implement HTMLImageElement.decode(), which the viewer
    // awaits before swapping artwork in.
    Object.defineProperty(HTMLImageElement.prototype, 'decode', {
      value: () => Promise.resolve(),
      configurable: true,
    });
    const image: ImageNode = {
      id: 'picture',
      type: 'image',
      transform,
      assetId: 'missing',
      fallbackPicture: {
        svg: `<svg xmlns="${SVG}" viewBox="0 0 4 4"><circle cx="2" cy="2" r="2"/></svg>`,
        paintBounds: { xPx: 0, yPx: 0, widthPx: 4, heightPx: 4, layoutWidthPx: 4, layoutHeightPx: 4 },
      },
    };
    const errors: string[] = [];
    const container = window.document.createElement('div');
    window.document.body.append(container);
    const viewer = new NormalizedPresentationViewer(container, presentation([image]), {
      onNodeError: (nodeId) => errors.push(nodeId),
    });
    await viewer.renderSlide(0);

    const img = container.querySelector<HTMLImageElement>('[data-rpv-node-id="picture"] img')!;
    expect(img.getAttribute('src')).toMatch(/^data:image\/svg\+xml/);
    expect(decodeURIComponent(img.getAttribute('src')!)).toContain('<circle');
    expect(errors).toEqual([]);
    viewer.destroy();
  });
});
