import type {
  PresentationDocument,
  PresentationEmbeddedFont,
  PresentationWarning,
} from '@extend-ai/react-pptx-model';
import type { PptxFontOptions, PptxFontSource } from './types';

const GENERIC_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-serif',
  'ui-sans-serif',
  'ui-monospace',
  'emoji',
  'math',
  'fangsong',
]);

/** Cross-platform, metrically similar substitutions for common Office typefaces. */
export const OFFICE_FONT_FALLBACKS: Readonly<Record<string, readonly string[]>> = {
  aptos: ['Aptos', 'Segoe UI', 'Arial', 'sans-serif'],
  'aptos display': ['Aptos Display', 'Segoe UI', 'Arial', 'sans-serif'],
  calibri: ['Calibri', 'Carlito', 'Arial', 'sans-serif'],
  'calibri light': ['Calibri Light', 'Carlito', 'Arial', 'sans-serif'],
  cambria: ['Cambria', 'Caladea', 'Georgia', 'serif'],
  'cambria math': ['Cambria Math', 'STIX Two Math', 'STIXGeneral', 'math'],
  arial: ['Arial', 'Liberation Sans', 'Helvetica', 'sans-serif'],
  'arial narrow': ['Arial Narrow', 'Liberation Sans Narrow', 'Arial', 'sans-serif'],
  'arial unicode ms': ['Arial Unicode MS', 'Noto Sans', 'Segoe UI Symbol', 'sans-serif'],
  'times new roman': ['Times New Roman', 'Liberation Serif', 'Times', 'serif'],
  'courier new': ['Courier New', 'Liberation Mono', 'Courier', 'monospace'],
  consolas: ['Consolas', 'Cascadia Mono', 'Liberation Mono', 'monospace'],
  'century gothic': ['Century Gothic', 'Avenir Next', 'Futura', 'Arial', 'sans-serif'],
  'franklin gothic medium': ['Franklin Gothic Medium', 'Arial', 'sans-serif'],
  'segoe ui': ['Segoe UI', 'Inter', 'Arial', 'sans-serif'],
  'sitka heading': ['Sitka Heading', 'Georgia', 'Times New Roman', 'serif'],
  sitka: ['Sitka', 'Georgia', 'Times New Roman', 'serif'],
  'source sans pro': ['Source Sans Pro', 'Helvetica Neue', 'Arial', 'sans-serif'],
  candara: ['Candara', 'Noto Sans', 'Arial', 'sans-serif'],
  constantia: ['Constantia', 'Georgia', 'serif'],
  corbel: ['Corbel', 'Noto Sans', 'Arial', 'sans-serif'],
  garamond: ['Garamond', 'EB Garamond', 'Georgia', 'serif'],
  'book antiqua': ['Book Antiqua', 'Palatino Linotype', 'Palatino', 'serif'],
  'palatino linotype': ['Palatino Linotype', 'Palatino', 'Georgia', 'serif'],
  'ms gothic': ['MS Gothic', 'Yu Gothic', 'Meiryo', 'Noto Sans CJK JP', 'sans-serif'],
  'ms pgothic': ['MS PGothic', 'Yu Gothic', 'Meiryo', 'Noto Sans CJK JP', 'sans-serif'],
  'ms mincho': ['MS Mincho', 'Yu Mincho', 'Noto Serif CJK JP', 'serif'],
  'ms pmincho': ['MS PMincho', 'Yu Mincho', 'Noto Serif CJK JP', 'serif'],
  meiryo: ['Meiryo', 'Yu Gothic', 'Noto Sans CJK JP', 'sans-serif'],
  'yu gothic': ['Yu Gothic', 'Meiryo', 'Noto Sans CJK JP', 'sans-serif'],
  'yu mincho': ['Yu Mincho', 'Noto Serif CJK JP', 'serif'],
  simsun: ['SimSun', 'Songti SC', 'Noto Serif CJK SC', 'serif'],
  simhei: ['SimHei', 'Heiti SC', 'Microsoft YaHei', 'Noto Sans CJK SC', 'sans-serif'],
  'microsoft yahei': ['Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'sans-serif'],
  dengxian: ['DengXian', 'Microsoft YaHei', 'Noto Sans CJK SC', 'sans-serif'],
  fangsong: ['FangSong', 'STFangsong', 'Noto Serif CJK SC', 'serif'],
  kaiti: ['KaiTi', 'STKaiti', 'Noto Serif CJK SC', 'serif'],
  mingliu: ['MingLiU', 'PingFang TC', 'Noto Serif CJK TC', 'serif'],
  pmingliu: ['PMingLiU', 'PingFang TC', 'Noto Serif CJK TC', 'serif'],
  'microsoft jhenghei': ['Microsoft JhengHei', 'PingFang TC', 'Noto Sans CJK TC', 'sans-serif'],
  'malgun gothic': ['Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans CJK KR', 'sans-serif'],
  batang: ['Batang', 'AppleMyungjo', 'Noto Serif CJK KR', 'serif'],
  gulim: ['Gulim', 'Apple SD Gothic Neo', 'Noto Sans CJK KR', 'sans-serif'],
  'nirmala ui': ['Nirmala UI', 'Noto Sans Devanagari', 'Noto Sans', 'sans-serif'],
  mangal: ['Mangal', 'Noto Sans Devanagari', 'Nirmala UI', 'sans-serif'],
  'traditional arabic': ['Traditional Arabic', 'Noto Naskh Arabic', 'Arial', 'serif'],
  'simplified arabic': ['Simplified Arabic', 'Noto Naskh Arabic', 'Arial', 'sans-serif'],
  'sakkal majalla': ['Sakkal Majalla', 'Noto Naskh Arabic', 'serif'],
  gisha: ['Gisha', 'Noto Sans Hebrew', 'Arial Hebrew', 'sans-serif'],
  david: ['David', 'Noto Serif Hebrew', 'Times New Roman', 'serif'],
};

const UNIVERSAL_SANS_FALLBACKS = [
  'Noto Sans',
  'Noto Sans CJK SC',
  'Noto Sans CJK JP',
  'Noto Sans CJK KR',
  'Noto Sans Arabic',
  'Noto Sans Hebrew',
  'Segoe UI Symbol',
  'sans-serif',
] as const;

function normalizeFamily(value: string): string {
  return value.trim().replace(/^['"]|['"]$/g, '');
}

function parseFamilyList(value: string): string[] {
  const result: string[] = [];
  let token = '';
  let quote = '';
  for (const character of value) {
    if ((character === '"' || character === "'") && (!quote || quote === character)) {
      quote = quote ? '' : character;
      token += character;
    } else if (character === ',' && !quote) {
      if (normalizeFamily(token)) result.push(normalizeFamily(token));
      token = '';
    } else token += character;
  }
  if (normalizeFamily(token)) result.push(normalizeFamily(token));
  return result;
}

function serializeFamily(family: string): string {
  if (GENERIC_FAMILIES.has(family.toLowerCase())) return family;
  return `"${family.replaceAll('"', '\\"')}"`;
}

function uniqueFamilies(families: readonly string[]): string[] {
  const seen = new Set<string>();
  return families.filter((family) => {
    const key = normalizeFamily(family).toLowerCase();
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function resolvePptxFontFamily(
  fontFamily: string,
  options: Pick<PptxFontOptions, 'fallbacks' | 'fallbackFamilies' | 'useOfficeFallbacks'> = {},
): string {
  const requested = parseFamilyList(fontFamily);
  const primary = requested[0] ?? '';
  const configured = options.fallbacks?.[primary] ?? options.fallbacks?.[primary.toLowerCase()];
  const configuredList = typeof configured === 'string' ? [configured] : (configured ?? []);
  const office =
    options.useOfficeFallbacks === false
      ? []
      : (OFFICE_FONT_FALLBACKS[primary.toLowerCase()] ?? []);
  const lower = primary.toLowerCase();
  const generic = /serif|times|cambria|mincho|ming|song|batang|david|garamond|palatino/.test(lower)
    ? ['Noto Serif', 'Noto Serif CJK SC', 'serif']
    : [...UNIVERSAL_SANS_FALLBACKS];
  return uniqueFamilies([
    ...(primary ? [primary] : []),
    ...configuredList,
    ...office,
    ...requested.slice(1),
    ...(options.fallbackFamilies ?? []),
    ...generic,
  ])
    .map(serializeFamily)
    .join(', ');
}

function fontSourceValue(source: string): string {
  if (/^(?:url|local)\(/i.test(source.trim())) return source;
  return `url(${JSON.stringify(source)})`;
}

function browserHasFont(family: string): boolean {
  if (
    typeof document.fonts?.check === 'function' &&
    !document.fonts.check(`12px ${serializeFamily(family)}`)
  ) {
    return false;
  }
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  if (!context) return true;
  const sample = 'mmmmmmmmmmlliWWW0123456789';
  return ['monospace', 'serif', 'sans-serif'].some((baseline) => {
    context.font = `72px ${baseline}`;
    const baselineWidth = context.measureText(sample).width;
    context.font = `72px ${serializeFamily(family)}, ${baseline}`;
    return Math.abs(context.measureText(sample).width - baselineWidth) > 0.01;
  });
}

function timeout<T>(promise: Promise<T>, milliseconds: number): Promise<T> {
  if (!Number.isFinite(milliseconds) || milliseconds <= 0) return promise;
  return new Promise<T>((resolve, reject) => {
    const timer = globalThis.setTimeout(
      () => reject(new Error('Font loading timed out.')),
      milliseconds,
    );
    void promise.then(
      (value) => {
        globalThis.clearTimeout(timer);
        resolve(value);
      },
      (error: unknown) => {
        globalThis.clearTimeout(timer);
        reject(error);
      },
    );
  });
}

export class PptxFontManager {
  private readonly faces: FontFace[] = [];
  private readonly objectUrls = new Set<string>();
  private readonly warned = new Set<string>();
  private readonly loadedFamilies = new Set<string>();

  constructor(
    private readonly options: PptxFontOptions = {},
    private readonly reportWarning: (warning: PresentationWarning) => void = () => {},
  ) {}

  private warn(warning: PresentationWarning): void {
    const key = `${warning.code}:${warning.feature ?? warning.message}`;
    if (this.warned.has(key)) return;
    this.warned.add(key);
    this.reportWarning(warning);
  }

  private embeddedSources(presentation: PresentationDocument): PptxFontSource[] {
    if (this.options.loadEmbeddedFonts === false) return [];
    return (presentation.embeddedFonts ?? []).flatMap((font: PresentationEmbeddedFont) => {
      const asset = presentation.assets[font.assetId];
      if (!asset?.data) return [];
      return [
        {
          family: font.family,
          source: asset.data,
          descriptors: { style: font.style, weight: font.weight },
        },
      ];
    });
  }

  private async loadSource(source: PptxFontSource): Promise<void> {
    if (typeof FontFace === 'undefined' || !document.fonts) return;
    let data: string | BufferSource;
    if (typeof source.source === 'string') data = fontSourceValue(source.source);
    else if (source.source instanceof Blob) {
      const url = URL.createObjectURL(source.source);
      this.objectUrls.add(url);
      data = `url(${JSON.stringify(url)})`;
    } else if (source.source instanceof Uint8Array) {
      data = source.source.slice().buffer as ArrayBuffer;
    } else data = source.source;
    const face = new FontFace(source.family, data, source.descriptors);
    document.fonts.add(face);
    this.faces.push(face);
    await timeout(face.load(), this.options.loadTimeoutMs ?? 5_000);
    this.loadedFamilies.add(source.family.toLowerCase());
    this.options.onFontLoaded?.(source.family, face);
  }

  async prepare(presentation: PresentationDocument): Promise<void> {
    const sources = [...this.embeddedSources(presentation), ...(this.options.sources ?? [])];
    const results = await Promise.allSettled(sources.map((source) => this.loadSource(source)));
    results.forEach((result, index) => {
      if (result.status === 'fulfilled') return;
      const family = sources[index]?.family ?? 'unknown';
      this.warn({
        code: 'missing-font',
        severity: 'warning',
        feature: family,
        message: `The supplied font “${family}” could not be loaded; a compatible fallback will be used.`,
      });
    });
    if (document.fonts && this.options.waitForFonts !== false) {
      await timeout(
        Promise.resolve(document.fonts.ready).then(() => undefined),
        this.options.loadTimeoutMs ?? 5_000,
      ).catch(() => undefined);
    }
    this.options.onFontsReady?.();
  }

  applyTo(root: HTMLElement, slideIndex?: number): void {
    const elements: Element[] = [
      root,
      ...root.querySelectorAll('[style*="font-family"], [font-family]'),
    ];
    for (const element of elements) {
      const raw =
        (element instanceof HTMLElement || element instanceof SVGElement
          ? element.style.fontFamily
          : '') || element.getAttribute('font-family');
      if (!raw) continue;
      const primary = parseFamilyList(raw)[0];
      if (!primary || primary.startsWith('+') || GENERIC_FAMILIES.has(primary.toLowerCase()))
        continue;
      if (element instanceof HTMLElement || element instanceof SVGElement) {
        element.style.fontFamily = resolvePptxFontFamily(raw, this.options);
      }
      if (
        this.options.reportMissingFonts === false ||
        this.loadedFamilies.has(primary.toLowerCase()) ||
        browserHasFont(primary)
      ) {
        continue;
      }
      this.warn({
        code: 'missing-font',
        severity: 'warning',
        feature: primary,
        ...(slideIndex === undefined ? {} : { slideIndex }),
        message: `Font “${primary}” is not available in this browser; the configured fallback stack is being used.`,
      });
    }
  }

  destroy(): void {
    if (document.fonts) for (const face of this.faces) document.fonts.delete(face);
    this.faces.length = 0;
    for (const url of this.objectUrls) URL.revokeObjectURL(url);
    this.objectUrls.clear();
  }
}
