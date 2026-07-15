import type {
  PresentationDocument,
  PresentationSearchResult,
  ShapeNode,
} from '@extend-ai/react-pptx-model';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { NormalizedPresentationViewer } from './normalized-viewer';

const presentation: PresentationDocument = {
  format: 'pptx',
  size: { widthEmu: 9_144_000, heightEmu: 6_858_000 },
  slides: [{ id: 'slide-1', index: 0, nodes: [] }],
  masters: [],
  layouts: [],
  themes: [],
  assets: {},
  warnings: [],
};

const transform = { x: 0, y: 0, width: 1_000_000, height: 500_000 };

function shape(
  id: string,
  text: string,
  options: Partial<Omit<ShapeNode, 'id' | 'type' | 'transform' | 'geometry' | 'paragraphs'>> = {},
): ShapeNode {
  return {
    id,
    type: 'shape',
    transform,
    geometry: { preset: 'rect' },
    paragraphs: [{ runs: [{ text }] }],
    ...options,
  };
}

function presentationWithSlides(count: number, nodes: ShapeNode[] = []): PresentationDocument {
  return {
    ...presentation,
    slides: Array.from({ length: count }, (_, index) => ({
      id: `slide-${index + 1}`,
      index,
      nodes: nodes.map((node) => ({ ...node })),
    })),
  };
}

function installObserverStub(): IntersectionObserverCallback[] {
  const callbacks: IntersectionObserverCallback[] = [];
  class ObserverStub {
    constructor(callback: IntersectionObserverCallback) {
      callbacks.push(callback);
    }
    observe() {}
    disconnect() {}
    unobserve() {}
    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }
    readonly root = null;
    readonly rootMargin = '';
    readonly thresholds = [];
  }
  vi.stubGlobal('IntersectionObserver', ObserverStub);
  return callbacks;
}

describe('normalized viewer scrolling', () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('uses a host-owned custom scroll element as the observer root', async () => {
    const observedRoots: Array<Element | Document | null> = [];
    class ObserverStub {
      constructor(_callback: IntersectionObserverCallback, options?: IntersectionObserverInit) {
        observedRoots.push(options?.root ?? null);
      }
      observe() {}
      disconnect() {}
      unobserve() {}
      takeRecords(): IntersectionObserverEntry[] {
        return [];
      }
      readonly root = null;
      readonly rootMargin = '';
      readonly thresholds = [];
    }
    vi.stubGlobal('IntersectionObserver', ObserverStub);
    const hostScrollArea = document.createElement('div');
    const container = document.createElement('div');
    document.body.append(hostScrollArea, container);
    const viewer = new NormalizedPresentationViewer(container, presentation);

    await viewer.renderList({ enabled: true, scrollElement: hostScrollArea });

    expect(observedRoots).toEqual([hostScrollArea, hostScrollArea]);
    viewer.destroy();
    hostScrollArea.remove();
    container.remove();
  });

  it('uses its viewport by default and reports the most visible slide', async () => {
    const callbacks: IntersectionObserverCallback[] = [];
    const observerRoots: Array<Element | Document | null | undefined> = [];
    class ObserverStub {
      constructor(nextCallback: IntersectionObserverCallback, options?: IntersectionObserverInit) {
        callbacks.push(nextCallback);
        observerRoots.push(options?.root);
      }
      observe() {}
      disconnect() {}
      unobserve() {}
      takeRecords(): IntersectionObserverEntry[] {
        return [];
      }
      readonly root = null;
      readonly rootMargin = '';
      readonly thresholds = [];
    }
    vi.stubGlobal('IntersectionObserver', ObserverStub);
    const container = document.createElement('div');
    container.scrollTop = 500;
    document.body.append(container);
    const threeSlides: PresentationDocument = {
      ...presentation,
      slides: [0, 1, 2].map((index) => ({ id: `slide-${index + 1}`, index, nodes: [] })),
    };
    const changes: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, threeSlides, {
      onSlideChange: (index) => changes.push(index),
    });

    await viewer.renderList({ enabled: true, initialSlides: 3 });

    expect(observerRoots).toEqual([container, container]);
    expect(container.scrollTop).toBe(0);
    const visibilityCallback = callbacks[1];
    expect(visibilityCallback).toBeDefined();
    const items = [...container.querySelectorAll<HTMLElement>('[data-rpv-list-item]')];
    visibilityCallback?.(
      [
        {
          target: items[0]!,
          isIntersecting: true,
          intersectionRatio: 0.8,
          boundingClientRect: { top: 20 },
        },
        {
          target: items[1]!,
          isIntersecting: true,
          intersectionRatio: 0.25,
          boundingClientRect: { top: 700 },
        },
      ] as unknown as IntersectionObserverEntry[],
      {} as IntersectionObserver,
    );
    expect(viewer.currentSlideIndex).toBe(0);

    visibilityCallback?.(
      [
        {
          target: items[0]!,
          isIntersecting: true,
          intersectionRatio: 0.1,
          boundingClientRect: { top: -600 },
        },
        {
          target: items[1]!,
          isIntersecting: true,
          intersectionRatio: 0.9,
          boundingClientRect: { top: 40 },
        },
      ] as unknown as IntersectionObserverEntry[],
      {} as IntersectionObserver,
    );
    expect(viewer.currentSlideIndex).toBe(1);
    expect(changes).toEqual([0, 1]);
    viewer.destroy();
    container.remove();
  });

  it('reports slide unmounts when replacing and destroying normalized slides', async () => {
    const container = document.createElement('div');
    const unmounted: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, presentation, {
      onSlideUnmounted: (index) => unmounted.push(index),
    });

    await viewer.renderSlide(0);
    await viewer.renderSlide(0);
    expect(unmounted).toEqual([0]);

    viewer.destroy();
    expect(unmounted).toEqual([0, 0]);
  });

  it('disposes thumbnail handles once and also cleans them up on destroy', () => {
    const container = document.createElement('div');
    const firstTarget = document.createElement('div');
    const secondTarget = document.createElement('div');
    const unmounted: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, presentation, {
      onSlideUnmounted: (index) => unmounted.push(index),
    });

    const first = viewer.renderThumbnailToContainer(0, firstTarget);
    viewer.renderThumbnailToContainer(0, secondTarget);
    first.dispose();
    first.dispose();
    expect(unmounted).toEqual([0]);

    viewer.destroy();
    expect(unmounted).toEqual([0, 0]);
  });

  it('keeps continuous mode mounted when zoom or fit changes', async () => {
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, presentation);

    await viewer.renderList({ enabled: false });
    expect(container.querySelectorAll('[data-rpv-list-item]')).toHaveLength(1);

    await viewer.setZoom(125);
    expect(container.querySelectorAll('[data-rpv-list-item]')).toHaveLength(1);
    expect(container.querySelector<HTMLElement>('[data-rpv-slide-wrapper]')?.style.width).toBe(
      '1200px',
    );

    await viewer.setFitMode('none');
    expect(container.querySelectorAll('[data-rpv-list-item]')).toHaveLength(1);
    viewer.destroy();
  });
});

describe('normalized viewer safety and fidelity', () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('exposes hidden-slide metadata on normalized slide sections', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        { id: 'hidden-slide', index: 0, hidden: true, nodes: [] },
        { id: 'visible-slide', index: 1, nodes: [] },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);
    expect(
      container.querySelector<HTMLElement>('[data-rpv-slide-index="0"]')?.dataset.rpvSlideHidden,
    ).toBe('true');

    await viewer.renderSlide(1);
    expect(
      container.querySelector<HTMLElement>('[data-rpv-slide-index="1"]')?.dataset.rpvSlideHidden,
    ).toBe('false');
    viewer.destroy();
  });

  it('renders only allowlisted model hyperlinks as anchors', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            shape('safe-node', 'Safe node', { hyperlink: 'https://example.com/path' }),
            shape('unsafe-node', 'Unsafe node', { hyperlink: 'javascript:alert(1)' }),
            {
              ...shape('run-links', ''),
              paragraphs: [
                {
                  runs: [
                    { text: 'Relative', hyperlink: '../guide' },
                    { text: 'Hash', hyperlink: '#section' },
                    { text: 'Mail', hyperlink: 'mailto:test@example.com' },
                    { text: 'Unsafe', hyperlink: 'data:text/html,bad' },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    expect(container.querySelector('[data-rpv-node-id="safe-node"]')).toBeInstanceOf(
      HTMLAnchorElement,
    );
    expect(container.querySelector('[data-rpv-node-id="unsafe-node"]')).not.toBeInstanceOf(
      HTMLAnchorElement,
    );
    const links = [
      ...container.querySelectorAll<HTMLAnchorElement>('[data-rpv-node-id="run-links"] a'),
    ];
    expect(links.map((link) => link.getAttribute('href'))).toEqual([
      '../guide',
      '#section',
      'mailto:test@example.com',
    ]);
    expect(
      [...container.querySelectorAll('[data-rpv-node-id="run-links"] span')].some(
        (node) => node.textContent === 'Unsafe',
      ),
    ).toBe(true);
    expect(container.innerHTML).not.toContain('javascript:');
    expect(container.innerHTML).not.toContain('data:text/html');
    viewer.destroy();
  });

  it('uses natural EMU dimensions and rejects model-provided CSS injection', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      size: { widthEmu: 4_762_500, heightEmu: 2_857_500 },
      slides: [
        {
          id: 'slide-1',
          index: 0,
          background: { type: 'solid', color: { value: 'url(https://evil.example/x)' } },
          nodes: [
            {
              ...shape('malicious-shape', 'Unsafe color'),
              fill: { type: 'solid', color: { value: 'var(--attacker)' } },
              textInsets: { top: 95_250, right: 190_500, bottom: 285_750, left: 381_000 },
              paragraphs: [
                { runs: [{ text: 'Unsafe color', color: { value: 'url(javascript:bad)' } }] },
              ],
            },
            {
              id: 'chart',
              type: 'chart',
              transform: { x: 1_000_000, y: 0, width: 1_000_000, height: 500_000 },
              chartType: 'bar',
              series: [{ values: [2], color: { value: 'var(--attacker)' } }],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const slide = container.querySelector<HTMLElement>('[data-rpv-slide-index="0"]')!;
    const maliciousShape = container.querySelector<HTMLElement>(
      '[data-rpv-node-id="malicious-shape"]',
    )!;
    const maliciousTextBody = maliciousShape.querySelector<HTMLElement>('[data-rpv-text-body]')!;
    expect(slide.style.width).toBe('500px');
    expect(slide.style.height).toBe('300px');
    expect(maliciousTextBody.style.padding).toBe('10px 20px 30px 40px');
    expect(`${slide.style.background} ${maliciousShape.style.background}`).not.toMatch(
      /evil|url|var/i,
    );
    expect(
      container.querySelector<HTMLElement>('[data-rpv-node-id="malicious-shape"] span')?.style
        .color,
    ).toBe('');
    expect(container.querySelector('rect')?.getAttribute('fill')).toBe('#4472c4');
    expect(container.innerHTML).not.toMatch(/attacker|javascript:bad|evil\.example/i);
    viewer.destroy();
  });

  it('maps DrawingML text metrics, wrapping, insets, and normal autofit to CSS', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              ...shape('text-layout', ''),
              autofit: { mode: 'normal', fontScale: 0.8, lineSpacingReduction: 0.1 },
              textWrap: 'none',
              spaceFirstLastParagraph: false,
              paragraphs: [
                {
                  runs: [
                    {
                      text: 'Scaled text',
                      fontSizePt: 20,
                      characterSpacingPt: 1.5,
                    },
                  ],
                  lineSpacing: { unit: 'percent', value: 1 },
                  spaceBefore: { unit: 'points', value: 9 },
                  marginLeftEmu: 95_250,
                  indentEmu: -47_625,
                },
                {
                  runs: [{ text: 'Exact spacing' }],
                  lineSpacing: { unit: 'points', value: 20 },
                },
                {
                  runs: [{ text: 'Reduced default spacing' }],
                  spaceAfter: { unit: 'points', value: 7 },
                },
              ],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const body = container.querySelector<HTMLElement>('[data-rpv-text-body]')!;
    const content = container.querySelector<HTMLElement>('[data-rpv-text-content]')!;
    const textNode = container.querySelector<HTMLElement>('[data-rpv-node-id="text-layout"]')!;
    const paragraphs = [...container.querySelectorAll<HTMLElement>('[data-rpv-text-paragraph]')];
    const scaledRun = paragraphs[0]!.querySelector<HTMLElement>('span')!;
    expect(body.style.padding).toBe('4.8px 9.6px');
    expect(textNode.style.overflowX).toBe('visible');
    expect(textNode.style.overflowY).toBe('visible');
    expect(content.style.whiteSpace).toBe('pre');
    expect(content.style.overflowWrap).toBe('normal');
    expect(paragraphs[0]!.style.lineHeight).toBe('1.035');
    expect(paragraphs[0]!.style.marginTop).toBe('');
    expect(paragraphs[0]!.style.marginLeft).toBe('10px');
    expect(paragraphs[0]!.style.textIndent).toBe('-5px');
    expect(paragraphs[1]!.style.lineHeight).toBe('20pt');
    expect(paragraphs[2]!.style.lineHeight).toBe('1.035');
    expect(paragraphs[2]!.style.marginBottom).toBe('');
    expect(scaledRun.style.fontSize).toBe('16pt');
    expect(scaledRun.style.letterSpacing).toBe('1.5pt');
    viewer.destroy();
  });

  it('grows shape-autofit text boxes and keeps body rotation independent from vertical flow', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              ...shape('growing-text', ''),
              autofit: { mode: 'shape' },
              textRotation: 90,
              paragraphs: [
                {
                  bullet: { kind: 'character', value: '•', sizePercent: 0.5 },
                  runs: [{ text: 'Growing text', fontSizePt: 20 }],
                },
              ],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const node = container.querySelector<HTMLElement>('[data-rpv-node-id="growing-text"]')!;
    const body = node.querySelector<HTMLElement>('[data-rpv-text-body]')!;
    const bullet = node.querySelector<HTMLElement>('[data-rpv-bullet]')!;
    expect(node.style.height).toBe('auto');
    expect(node.style.minHeight).toBe(`${(transform.height / presentation.size.heightEmu) * 100}%`);
    expect(body.style.position).toBe('relative');
    expect(body.style.transform).toBe('rotate(90deg)');
    expect(body.style.writingMode).toBe('');
    expect(bullet.style.fontSize).toBe('10pt');
    viewer.destroy();
  });

  it('converts gradient vectors and sorts DrawingML stops by position', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              ...shape('gradient', ''),
              fill: {
                type: 'gradient',
                angle: 90,
                stops: [
                  { position: 1, color: { value: '#ffffff' } },
                  { position: 0, color: { value: '#000000' } },
                ],
              },
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const background = container.querySelector<HTMLElement>('[data-rpv-node-id="gradient"]')!.style
      .background;
    expect(background).toContain('linear-gradient(180deg');
    expect(background).toMatch(
      /(?:#000000|rgb\(0, 0, 0\)) 0%.*(?:#ffffff|rgb\(255, 255, 255\)) 100%/i,
    );
    viewer.destroy();
  });

  it('isolates image-fill opacity and renders normalized custom geometry as SVG', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      assets: {
        picture: {
          id: 'picture',
          contentType: 'image/png',
          byteLength: 1,
          url: 'data:image/png;base64,AA==',
        },
      },
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              ...shape('transparent-fill', 'Foreground text'),
              fill: {
                type: 'image',
                assetId: 'picture',
                mode: 'stretch',
                opacity: 0.14,
              },
            },
            {
              id: 'custom-path',
              type: 'shape',
              transform: { ...transform, x: 1_500_000 },
              geometry: { path: 'M 0 0 L 1 0.5 L 0 1 Z' },
              fill: { type: 'solid', color: { value: '#ff0000' } },
              line: { color: { value: '#000000' }, width: 2 },
              paragraphs: [],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const fillLayer = container.querySelector<HTMLElement>(
      '[data-rpv-node-id="transparent-fill"] [data-rpv-image-fill]',
    )!;
    expect(fillLayer.style.opacity).toBe('0.14');
    expect(fillLayer.style.backgroundImage).toContain('data:image/png');
    const svg = container.querySelector<SVGSVGElement>(
      '[data-rpv-node-id="custom-path"] [data-rpv-custom-geometry]',
    )!;
    const path = svg.querySelector('path')!;
    expect(svg.getAttribute('viewBox')).toBe('0 0 1 1');
    expect(path.getAttribute('d')).toBe('M 0 0 L 1 0.5 L 0 1 Z');
    expect(path.style.fill).toMatch(/#ff0000|rgb\(255, 0, 0\)/i);
    expect(path.style.strokeWidth).toBe('2px');
    viewer.destroy();
  });

  it('clips directional arrow presets to their DrawingML silhouettes', async () => {
    const arrows: ShapeNode[] = ['rightArrow', 'leftArrow', 'upArrow', 'downArrow'].map(
      (preset, index) => ({
        ...shape(`arrow-${index}`, ''),
        geometry: { preset },
      }),
    );
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(
      container,
      presentationWithSlides(1, arrows),
    );

    await viewer.renderSlide(0);

    expect(container.querySelector<HTMLElement>('[data-rpv-node-id="arrow-0"]')?.style.clipPath).toBe(
      'polygon(0 25%, 75% 25%, 75% 0, 100% 50%, 75% 100%, 75% 75%, 0 75%)',
    );
    expect(container.querySelector<HTMLElement>('[data-rpv-node-id="arrow-1"]')?.style.clipPath).toContain(
      'polygon(',
    );
    expect(container.querySelector<HTMLElement>('[data-rpv-node-id="arrow-2"]')?.style.clipPath).toContain(
      'polygon(',
    );
    expect(container.querySelector<HTMLElement>('[data-rpv-node-id="arrow-3"]')?.style.clipPath).toContain(
      'polygon(',
    );
    viewer.destroy();
  });

  it('stretches picture fills and expands source-rectangle crops to the full frame', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      assets: {
        picture: {
          id: 'picture',
          contentType: 'image/png',
          byteLength: 1,
          url: 'data:image/png;base64,AA==',
        },
      },
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              id: 'plain-image',
              type: 'image',
              transform,
              assetId: 'picture',
              preserveAspectRatio: true,
            },
            {
              id: 'stretched-image',
              type: 'image',
              transform: { ...transform, x: 500_000 },
              assetId: 'picture',
              preserveAspectRatio: false,
            },
            {
              id: 'cropped-image',
              type: 'image',
              transform: { ...transform, x: 1_000_000 },
              assetId: 'picture',
              crop: { top: 0.1, right: 0.2, bottom: 0.3, left: 0.25 },
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const plain = container.querySelector<HTMLImageElement>(
      '[data-rpv-node-id="plain-image"] img',
    )!;
    const stretched = container.querySelector<HTMLImageElement>(
      '[data-rpv-node-id="stretched-image"] img',
    )!;
    const cropped = container.querySelector<HTMLImageElement>(
      '[data-rpv-node-id="cropped-image"] img',
    )!;
    expect(plain.style.objectFit).toBe('contain');
    expect(stretched.style.objectFit).toBe('fill');
    expect(cropped.style.objectFit).toBe('fill');
    expect(Number.parseFloat(cropped.style.width)).toBeCloseTo(181.818_181_8);
    expect(Number.parseFloat(cropped.style.height)).toBeCloseTo(166.666_666_7);
    expect(Number.parseFloat(cropped.style.left)).toBeCloseTo(-45.454_545_5);
    expect(Number.parseFloat(cropped.style.top)).toBeCloseTo(-16.666_666_7);
    viewer.destroy();
  });

  it('ignores sparse and non-finite chart values instead of emitting invalid SVG geometry', async () => {
    const unsafeValues = [
      1,
      undefined,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      -2,
    ] as unknown as Array<number | null>;
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              id: 'sparse-chart',
              type: 'chart',
              transform,
              chartType: 'bar',
              series: [{ values: unsafeValues }],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const rectangles = [
      ...container.querySelectorAll<SVGRectElement>('[data-rpv-node-id="sparse-chart"] rect'),
    ];
    expect(rectangles).toHaveLength(2);
    expect(rectangles.every((rectangle) => !rectangle.outerHTML.includes('NaN'))).toBe(true);
    expect(rectangles.map((rectangle) => rectangle.getAttribute('y'))).toEqual(['60', '220']);
    expect(rectangles.map((rectangle) => rectangle.getAttribute('height'))).toEqual(['160', '320']);
    viewer.destroy();
  });

  it('preserves the established scale floor for positive-only charts', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              id: 'positive-chart',
              type: 'chart',
              transform,
              chartType: 'bar',
              series: [{ values: [0.25, 0.5] }],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const rectangles = [
      ...container.querySelectorAll<SVGRectElement>('[data-rpv-node-id="positive-chart"] rect'),
    ];
    expect(rectangles.map((rectangle) => rectangle.getAttribute('y'))).toEqual(['425', '310']);
    expect(rectangles.map((rectangle) => rectangle.getAttribute('height'))).toEqual(['115', '230']);
    viewer.destroy();
  });

  it('applies table-cell text margins, vertical alignment, and rotation', async () => {
    const documentModel: PresentationDocument = {
      ...presentation,
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [
            {
              id: 'table-layout',
              type: 'table',
              transform,
              rowHeights: [500_000],
              rows: [
                [
                  {
                    verticalAlignment: 'bottom',
                    textRotation: 270,
                    textInsets: {
                      top: 95_250,
                      right: 190_500,
                      bottom: 285_750,
                      left: 381_000,
                    },
                    paragraphs: [{ runs: [{ text: 'Rotated cell' }] }],
                  },
                ],
              ],
            },
          ],
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    const cell = container.querySelector<HTMLTableCellElement>('td')!;
    const body = cell.querySelector<HTMLElement>('[data-rpv-table-text-body]')!;
    expect(cell.style.verticalAlign).toBe('bottom');
    expect(cell.style.padding).toBe('0px');
    expect(cell.style.border).toBe('');
    expect(body.style.padding).toBe('10px 20px 30px 40px');
    expect(body.style.position).toBe('absolute');
    expect(body.style.justifyContent).toBe('flex-end');
    expect(body.style.writingMode).toBe('vertical-rl');
    expect(body.style.transform).toBe('rotate(180deg)');
    viewer.destroy();
  });

  it('caches binary asset URLs by asset id and revokes them on destroy', async () => {
    const createObjectURL = vi.fn(() => 'blob:shared');
    const revokeObjectURL = vi.fn();
    vi.stubGlobal('URL', { createObjectURL, revokeObjectURL });
    const documentModel: PresentationDocument = {
      ...presentation,
      assets: {
        shared: {
          id: 'shared',
          contentType: 'image/png',
          byteLength: 3,
          data: new Uint8Array([1, 2, 3]),
        },
      },
      slides: [
        {
          id: 'slide-1',
          index: 0,
          nodes: [0, 1].map((index) => ({
            id: `image-${index}`,
            type: 'image' as const,
            transform: { ...transform, x: index * 1_000_000 },
            assetId: 'shared',
          })),
        },
      ],
    };
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, documentModel);

    await viewer.renderSlide(0);

    expect(createObjectURL).toHaveBeenCalledTimes(1);
    expect(
      [...container.querySelectorAll<HTMLImageElement>('img')].map((image) => image.src),
    ).toEqual(['blob:shared', 'blob:shared']);
    viewer.destroy();
    viewer.destroy();
    expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    expect(revokeObjectURL).toHaveBeenCalledWith('blob:shared');
  });
});

describe('normalized viewer generations and windowing', () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('unmounts slides outside overscan and mounts navigation targets', async () => {
    const callbacks = installObserverStub();
    const container = document.createElement('div');
    const changes: number[] = [];
    const unmounted: number[] = [];
    const viewer = new NormalizedPresentationViewer(container, presentationWithSlides(3), {
      onSlideChange: (index) => changes.push(index),
      onSlideUnmounted: (index) => unmounted.push(index),
    });

    await viewer.renderList({ enabled: true, initialSlides: 1 });

    const items = [...container.querySelectorAll<HTMLElement>('[data-rpv-list-item]')];
    expect(items[0]!.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
    expect(items[2]!.querySelector('[data-rpv-slide-wrapper]')).toBeNull();
    callbacks[0]?.(
      [
        { target: items[0], isIntersecting: false },
        { target: items[2], isIntersecting: true },
      ] as unknown as IntersectionObserverEntry[],
      {} as IntersectionObserver,
    );
    await Promise.resolve();
    expect(items[0]!.querySelector('[data-rpv-slide-wrapper]')).toBeNull();
    expect(items[2]!.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
    expect(unmounted).toContain(0);

    await viewer.goToSlide(1);
    expect(items[1]!.querySelector('[data-rpv-slide-wrapper]')).not.toBeNull();
    expect(changes).toEqual([0, 1]);
    viewer.destroy();
  });

  it('mounts and scopes highlights to the result slide as a visible sibling overlay', async () => {
    installObserverStub();
    const repeatedNode = shape('repeated-id', 'Repeated');
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(
      container,
      presentationWithSlides(2, [repeatedNode]),
    );
    const result: PresentationSearchResult = {
      slideIndex: 1,
      nodeId: 'repeated-id',
      nodeType: 'shape',
      text: 'Repeated',
      matchStart: 0,
      matchEnd: 8,
      snippet: 'Repeated',
    };

    await viewer.renderList({ enabled: true, initialSlides: 1 });
    await viewer.highlightSearchResult(result, { scrollIntoView: false });

    const firstSlide = container.querySelector<HTMLElement>('[data-rpv-slide-index="0"]')!;
    const secondSlide = container.querySelector<HTMLElement>('[data-rpv-slide-index="1"]')!;
    const secondNode = secondSlide.querySelector<HTMLElement>('[data-rpv-node-id="repeated-id"]')!;
    const highlight = secondSlide.querySelector<HTMLElement>('.rpv-search-highlight')!;
    expect(firstSlide.querySelector('.rpv-search-highlight')).toBeNull();
    expect(highlight).not.toBeNull();
    expect(secondNode.contains(highlight)).toBe(false);
    expect(highlight.parentElement).toBe(secondNode.parentElement);
    expect(highlight.style.outline).toContain('#ef8b2c');
    expect(secondNode.style.overflowX).toBe('visible');
    expect(secondNode.style.overflowY).toBe('visible');
    viewer.destroy();
  });

  it('suppresses stale render completion and callbacks after destroy', async () => {
    const changes: number[] = [];
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, presentationWithSlides(2), {
      onSlideChange: (index) => changes.push(index),
    });

    const first = viewer.renderSlide(0);
    const second = viewer.renderSlide(1);
    await Promise.all([first, second]);
    expect(changes).toEqual([1]);

    changes.length = 0;
    const pending = viewer.renderSlide(0);
    viewer.destroy();
    await pending;
    expect(changes).toEqual([]);
    expect(container.childElementCount).toBe(0);
  });

  it('renders an empty presentation without throwing or reporting a slide change', async () => {
    const container = document.createElement('div');
    const onSlideChange = vi.fn();
    const viewer = new NormalizedPresentationViewer(container, presentationWithSlides(0), {
      onSlideChange,
    });

    await expect(viewer.renderSlide()).resolves.toBeUndefined();
    await expect(viewer.renderList()).resolves.toBeUndefined();
    expect(onSlideChange).not.toHaveBeenCalled();
    expect(container.childElementCount).toBe(0);
    viewer.destroy();
  });

  it('honors batchSize when mounting a non-windowed list', async () => {
    const animationFrame = vi
      .spyOn(globalThis, 'requestAnimationFrame')
      .mockImplementation((callback) => {
        callback(0);
        return 1;
      });
    const container = document.createElement('div');
    const viewer = new NormalizedPresentationViewer(container, presentationWithSlides(5));

    await viewer.renderList({ enabled: false, batchSize: 2 });

    expect(animationFrame).toHaveBeenCalledTimes(2);
    expect(container.querySelectorAll('[data-rpv-slide-wrapper]')).toHaveLength(5);
    viewer.destroy();
  });
});
