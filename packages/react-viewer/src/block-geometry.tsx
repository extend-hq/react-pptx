import * as React from 'react';
import { createPortal } from 'react-dom';
import type {
  PptxSourceRef,
  PresentationBlock,
  PresentationWord,
} from '@extend-ai/react-pptx-model';
import type { PptxViewerController } from './types';

export interface PptxBoundingBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}
export interface PptxBlockFragment {
  nodeId: string;
  documentId: string;
  slideIndex: number;
  slideWidth: number;
  slideHeight: number;
  boundingBox: PptxBoundingBox;
  coordinateSpace: 'slide-css-px';
}
export interface PptxBlockGeometry {
  status: 'rendered' | 'partially_rendered' | 'not_mounted' | 'not_visible' | 'unresolved';
  scope: 'mounted_slides';
  fragments: PptxBlockFragment[];
  issues: Array<{
    source: PptxSourceRef;
    reason:
      | 'document_mismatch'
      | 'source_mismatch'
      | 'unsupported_range'
      | 'range_text_mismatch'
      | 'not_on_slide'
      | 'not_mounted'
      | 'not_visible';
  }>;
}

function slides(root: HTMLElement): HTMLElement[] {
  return [
    ...(root.matches('[data-rpv-slide-index]') ? [root] : []),
    ...root.querySelectorAll<HTMLElement>('[data-rpv-slide-index]'),
  ].filter((s) => s.dataset.rpvCitationSurface !== 'thumbnail');
}

function clippedBox(
  element: HTMLElement,
  surface: HTMLElement,
  input: PptxBoundingBox,
): PptxBoundingBox | undefined {
  let box = { left: input.left, top: input.top, right: input.right, bottom: input.bottom };
  for (let parent: HTMLElement | null = element; parent; parent = parent.parentElement) {
    const style = getComputedStyle(parent);
    if (
      parent.hidden ||
      style.display === 'none' ||
      style.visibility === 'hidden' ||
      Number(style.opacity) === 0
    )
      return;
    if (parent !== element || parent === surface) {
      const bounds = parent.getBoundingClientRect();
      if (parent === surface || /hidden|clip|scroll|auto/.test(style.overflowX)) {
        box.left = Math.max(box.left, bounds.left);
        box.right = Math.min(box.right, bounds.right);
      }
      if (parent === surface || /hidden|clip|scroll|auto/.test(style.overflowY)) {
        box.top = Math.max(box.top, bounds.top);
        box.bottom = Math.min(box.bottom, bounds.bottom);
      }
    }
    if (parent === surface) break;
  }
  return box.right > box.left && box.bottom > box.top ? box : undefined;
}

function rangeRects(
  element: HTMLElement,
  source: PptxSourceRef,
  expected?: string,
): { rects: DOMRect[]; error?: 'unsupported_range' | 'range_text_mismatch' } {
  const text = element.dataset.rpvSourceText;
  if (text === undefined || source.offsetEncoding !== 'utf-16')
    return { rects: [], error: 'unsupported_range' };
  const start = source.start ?? 0,
    end = source.end ?? text.length;
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > text.length ||
    (source.start === undefined) !== (source.end === undefined)
  )
    return { rects: [], error: 'unsupported_range' };
  const splitsSurrogate = (at: number) =>
    at > 0 &&
    at < text.length &&
    /[\uD800-\uDBFF]/.test(text[at - 1]!) &&
    /[\uDC00-\uDFFF]/.test(text[at]!);
  if (splitsSurrogate(start) || splitsSurrogate(end))
    return { rects: [], error: 'unsupported_range' };
  if (expected !== undefined && text.slice(start, end) !== expected)
    return { rects: [], error: 'range_text_mismatch' };
  const rects: DOMRect[] = [];
  const covered: Array<[number, number]> = [];
  for (const span of element.querySelectorAll<HTMLElement>('[data-rpv-source-start]')) {
    if (span.closest('[data-rpv-text-paragraph]') !== element) continue;
    const from = Number(span.dataset.rpvSourceStart),
      to = Number(span.dataset.rpvSourceEnd);
    const lo = Math.max(start, from),
      hi = Math.min(end, to);
    if (hi <= lo) continue;
    const original = span.dataset.rpvSourceText ?? '';
    if (original !== text.slice(from, to)) return { rects: [], error: 'range_text_mismatch' };
    if (span.dataset.rpvMath !== undefined) {
      if (source.start !== undefined) return { rects: [], error: 'unsupported_range' };
      rects.push(span.getBoundingClientRect());
      covered.push([lo, hi]);
      continue;
    }
    if (span.textContent !== original) return { rects: [], error: 'range_text_mismatch' };
    const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
    let offset = from;
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      const next = offset + node.length;
      if (Math.min(hi, next) > Math.max(lo, offset)) {
        const range = document.createRange();
        range.setStart(node, Math.max(lo, offset) - offset);
        range.setEnd(node, Math.min(hi, next) - offset);
        rects.push(...range.getClientRects());
      }
      offset = next;
    }
    covered.push([lo, hi]);
  }
  let cursor = start;
  for (const [lo, hi] of covered.sort((a, b) => a[0] - b[0])) {
    if (/[^\t\n]/.test(text.slice(cursor, lo))) return { rects: [], error: 'unsupported_range' };
    cursor = Math.max(cursor, hi);
  }
  if (/[^\t\n]/.test(text.slice(cursor, end))) return { rects: [], error: 'unsupported_range' };
  return { rects };
}

export function resolvePptxSourceGeometry(
  root: HTMLElement,
  refs: readonly PptxSourceRef[],
  expectedText?: string,
): PptxBlockGeometry {
  const fragments: PptxBlockFragment[] = [],
    issues: PptxBlockGeometry['issues'] = [];
  const surfaces = slides(root);
  for (const source of refs) {
    const issue = (reason: PptxBlockGeometry['issues'][number]['reason']) =>
      issues.push({ source, reason });
    if (source.story !== 'slide') {
      issue('not_on_slide');
      continue;
    }
    if (
      !source.documentId ||
      (surfaces.length && !surfaces.some((s) => s.dataset.rpvDocumentId === source.documentId))
    ) {
      issue('document_mismatch');
      continue;
    }
    const surface = surfaces.find(
      (s) =>
        s.dataset.rpvDocumentId === source.documentId &&
        Number(s.dataset.rpvSlideIndex) === source.slideIndex,
    );
    if (!surface) {
      issue('not_mounted');
      continue;
    }
    if (
      surface.dataset.rpvSlideId !== source.slideId ||
      !source.nodePath.length ||
      source.nodePath.some((n) => !Number.isInteger(n) || n < 0)
    ) {
      issue('source_mismatch');
      continue;
    }
    const node = [...surface.querySelectorAll<HTMLElement>('[data-rpv-node-path]')].find(
      (n) => n.dataset.rpvNodePath === source.nodePath.join('.'),
    );
    if (
      !node ||
      node.dataset.rpvNodeId !== source.nodeId ||
      (source.partName !== undefined &&
        (node.closest<HTMLElement>('[data-rpv-source-part]')?.dataset.rpvSourcePart ??
          surface.dataset.rpvSourcePart) !== source.partName) ||
      (source.sourceId !== undefined && node.dataset.rpvSourceId !== source.sourceId)
    ) {
      issue('source_mismatch');
      continue;
    }
    let target = node;
    if (source.rowIndex !== undefined || source.columnIndex !== undefined) {
      const cell = [...node.querySelectorAll<HTMLElement>('[data-rpv-cell-row]')].find(
        (c) =>
          c.closest('[data-rpv-node-path]') === node &&
          Number(c.dataset.rpvCellRow) === source.rowIndex &&
          Number(c.dataset.rpvCellColumn) === source.columnIndex,
      );
      if (!cell) {
        issue('source_mismatch');
        continue;
      }
      target = cell;
    }
    if (source.paragraphIndex !== undefined) {
      const paragraph = [...target.querySelectorAll<HTMLElement>('[data-rpv-text-paragraph]')].find(
        (p) =>
          p.closest('[data-rpv-node-path]') === node &&
          Number(p.dataset.rpvTextParagraph) === source.paragraphIndex &&
          (target === node
            ? !p.closest('[data-rpv-cell-row]')
            : p.closest('[data-rpv-cell-row]') === target),
      );
      if (!paragraph) {
        issue('source_mismatch');
        continue;
      }
      target = paragraph;
    } else if (source.start !== undefined || source.end !== undefined) {
      issue('unsupported_range');
      continue;
    }
    const measured =
      source.paragraphIndex !== undefined
        ? rangeRects(target, source, expectedText)
        : { rects: [target.getBoundingClientRect()] };
    if (measured.error) {
      issue(measured.error);
      continue;
    }
    const bounds = surface.getBoundingClientRect();
    const width = parseFloat(surface.style.width),
      height = parseFloat(surface.style.height);
    if (!(bounds.width > 0 && bounds.height > 0 && width > 0 && height > 0)) {
      issue('not_visible');
      continue;
    }
    let count = 0;
    for (const rect of measured.rects) {
      const box = clippedBox(target, surface, rect);
      if (!box) continue;
      fragments.push({
        documentId: source.documentId,
        nodeId: source.nodeId,
        slideIndex: source.slideIndex,
        slideWidth: width,
        slideHeight: height,
        coordinateSpace: 'slide-css-px',
        boundingBox: {
          left: ((box.left - bounds.left) * width) / bounds.width,
          right: ((box.right - bounds.left) * width) / bounds.width,
          top: ((box.top - bounds.top) * height) / bounds.height,
          bottom: ((box.bottom - bounds.top) * height) / bounds.height,
        },
      });
      count += 1;
    }
    if (!count) issue('not_visible');
  }
  const reasons = new Set(issues.map((i) => i.reason));
  return {
    scope: 'mounted_slides',
    fragments,
    issues,
    status: fragments.length
      ? issues.length
        ? 'partially_rendered'
        : 'rendered'
      : reasons.size === 1 && reasons.has('not_mounted')
        ? 'not_mounted'
        : reasons.size === 1 && reasons.has('not_visible')
          ? 'not_visible'
          : 'unresolved',
  };
}

export type PptxHighlightTarget = Pick<PresentationBlock, 'id' | 'content' | 'sourceRefs'>;
export function resolvePptxBlockGeometries(
  root: HTMLElement,
  blocks: readonly PptxHighlightTarget[],
): Map<string, PptxBlockGeometry> {
  return new Map(
    blocks.map((block) => [
      block.id,
      resolvePptxSourceGeometry(
        root,
        block.sourceRefs,
        block.sourceRefs.length === 1 && block.sourceRefs[0]?.paragraphIndex !== undefined
          ? block.content
          : undefined,
      ),
    ]),
  );
}
export function resolvePptxWordGeometries(
  root: HTMLElement,
  words: readonly PresentationWord[],
): Map<string, PptxBlockGeometry> {
  return resolvePptxBlockGeometries(root, words);
}
export function normalizedPptxBoundingBox(fragment: PptxBlockFragment): PptxBoundingBox {
  const b = fragment.boundingBox;
  return {
    left: b.left / fragment.slideWidth,
    top: b.top / fragment.slideHeight,
    right: b.right / fragment.slideWidth,
    bottom: b.bottom / fragment.slideHeight,
  };
}
export async function revealPptxSource(
  controller: PptxViewerController,
  root: HTMLElement,
  source: PptxSourceRef,
): Promise<PptxBlockGeometry> {
  if (controller.getDocument()?.documentId !== source.documentId)
    return {
      scope: 'mounted_slides',
      status: 'unresolved',
      fragments: [],
      issues: [{ source, reason: 'document_mismatch' }],
    };
  if (source.story !== 'slide')
    return {
      scope: 'mounted_slides',
      status: 'unresolved',
      fragments: [],
      issues: [{ source, reason: 'not_on_slide' }],
    };
  await controller.goToSlide(source.slideIndex);
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  return resolvePptxSourceGeometry(root, [source]);
}

export interface PptxBlockHighlightsProps {
  rootRef: React.RefObject<HTMLElement | null>;
  block?: PptxHighlightTarget;
  blocks?: readonly PptxHighlightTarget[];
  color?: string;
  colorForBlock?: (block: PptxHighlightTarget) => string;
  onGeometry?: (geometry: PptxBlockGeometry) => void;
  onGeometries?: (geometry: Map<string, PptxBlockGeometry>) => void;
}
export function PptxBlockHighlights({
  rootRef,
  block,
  blocks,
  color = '#2563eb',
  onGeometry,
  onGeometries,
  colorForBlock,
}: PptxBlockHighlightsProps): React.JSX.Element {
  const targets = React.useMemo(
    () => [
      ...new Map([...(blocks ?? []), ...(block ? [block] : [])].map((b) => [b.id, b])).values(),
    ],
    [block, blocks],
  );
  const [geometry, setGeometry] = React.useState(new Map<string, PptxBlockGeometry>());
  const callback = React.useRef({ onGeometry, onGeometries, block });
  callback.current = { onGeometry, onGeometries, block };
  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || !targets.length) {
      setGeometry(new Map());
      return;
    }
    let frame = 0,
      stopped = false;
    const update = () => {
      frame = 0;
      if (stopped) return;
      const next = resolvePptxBlockGeometries(root, targets);
      setGeometry((previous) =>
        JSON.stringify([...previous]) === JSON.stringify([...next]) ? previous : next,
      );
      callback.current.onGeometries?.(next);
      const active = callback.current.block && next.get(callback.current.block.id);
      if (active) callback.current.onGeometry?.(active);
    };
    const schedule = () => {
      if (!frame && !stopped) frame = requestAnimationFrame(update);
    };
    const resize = new ResizeObserver(schedule);
    resize.observe(root);
    const mutations = new MutationObserver((records) => {
      if (
        records.some(
          (record) =>
            !(
              record.target instanceof Element &&
              record.target.closest('[data-rpv-block-highlight]')
            ) &&
            (record.type !== 'childList' ||
              [...record.addedNodes, ...record.removedNodes].some(
                (n) => !(n instanceof Element && n.hasAttribute('data-rpv-block-highlight')),
              )),
        )
      ) {
        for (const surface of slides(root)) resize.observe(surface);
        schedule();
      }
    });
    mutations.observe(root, {
      subtree: true,
      childList: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['style', 'class', 'hidden', 'data-rpv-document-id'],
    });
    for (const surface of slides(root)) resize.observe(surface);
    document.fonts?.addEventListener('loadingdone', schedule);
    void document.fonts?.ready.then(schedule);
    window.addEventListener('resize', schedule);
    root.addEventListener('scroll', schedule, true);
    root.addEventListener('load', schedule, true);
    update();
    return () => {
      stopped = true;
      cancelAnimationFrame(frame);
      mutations.disconnect();
      resize.disconnect();
      document.fonts?.removeEventListener('loadingdone', schedule);
      window.removeEventListener('resize', schedule);
      root.removeEventListener('scroll', schedule, true);
      root.removeEventListener('load', schedule, true);
    };
  }, [rootRef, targets]);
  const surfaces = rootRef.current ? slides(rootRef.current) : [];
  return (
    <>
      {targets.flatMap((target) =>
        (geometry.get(target.id)?.fragments ?? []).map((fragment, index) => {
          const surface = surfaces.find(
            (s) =>
              s.dataset.rpvDocumentId === fragment.documentId &&
              Number(s.dataset.rpvSlideIndex) === fragment.slideIndex,
          );
          if (!surface) return null;
          const box = fragment.boundingBox;
          const tint = colorForBlock?.(target) ?? color;
          return createPortal(
            <div
              data-rpv-block-highlight=""
              aria-hidden="true"
              style={{
                position: 'absolute',
                pointerEvents: 'none',
                left: box.left,
                top: box.top,
                width: box.right - box.left,
                height: box.bottom - box.top,
                boxSizing: 'border-box',
                border: `${target.id === block?.id ? 2 : 1}px solid ${tint}`,
                background: `color-mix(in srgb, ${tint} 14%, transparent)`,
                zIndex: 2147483646,
              }}
            />,
            surface,
            `${target.id}/${index}`,
          );
        }),
      )}
    </>
  );
}

export interface PptxWordHighlightsProps {
  rootRef: React.RefObject<HTMLElement | null>;
  word?: PresentationWord;
  words?: readonly PresentationWord[];
  color?: string;
  onGeometry?: (geometry: PptxBlockGeometry) => void;
}
export function PptxWordHighlights({
  word,
  words,
  ...props
}: PptxWordHighlightsProps): React.JSX.Element {
  return (
    <PptxBlockHighlights
      {...props}
      {...(word ? { block: word } : {})}
      {...(words ? { blocks: words } : {})}
    />
  );
}
