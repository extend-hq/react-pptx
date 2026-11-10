import * as React from 'react';
import { createPortal } from 'react-dom';
import type { ParsedPresentation } from '@extend-ai/react-pptx';
import type { Rect, SlideNode } from '@extend-ai/react-pptx-model';
import type { ObjectEdit } from '@/lib/pptx-editing';
import { startInlineTextEdit, type InlineTextSession } from '@/lib/inline-text-editing';

const EMU = 914400;
type Selection = { slide: number; index: number };
type Props = {
  rootRef: React.RefObject<HTMLDivElement | null>;
  parsed: ParsedPresentation | null;
  disabled: boolean;
  busy: boolean;
  revision: number;
  slideIndex?: number;
  onEdit: (edit: ObjectEdit) => Promise<void>;
};

export function ObjectEditor({
  rootRef,
  parsed,
  disabled,
  busy,
  revision,
  slideIndex,
  onEdit,
}: Props) {
  const [selection, setSelection] = React.useState<Selection | null>(null);
  const [target, setTarget] = React.useState<HTMLElement | null>(null);
  const [preview, setPreview] = React.useState<Rect | null>(null);
  const inline = React.useRef<InlineTextSession | null>(null);
  const [textMode, setTextMode] = React.useState(false);
  const cancelDrag = React.useRef<(() => void) | null>(null);
  const node = selection && parsed?.document.slides[selection.slide]?.nodes[selection.index];
  const slide = target?.closest<HTMLElement>('[data-rpv-slide-index]');
  const select = React.useCallback(
    (element: HTMLElement) => {
      const surface = element.closest<HTMLElement>('[data-rpv-slide-index]');
      if (!surface) return;
      const index = Number(element.dataset.rpvNodePath);
      const slideIndex = Number(surface.dataset.rpvSlideIndex);
      const candidate = parsed?.document.slides[slideIndex]?.nodes[index];
      if (
        !candidate ||
        !candidate.sourceId ||
        candidate.sourcePart !== parsed?.document.slides[slideIndex]?.sourcePart ||
        candidate.editing?.locks?.noSelection ||
        (!candidate.editing && candidate.type !== 'group') ||
        candidate.editing?.kind === 'contentPart'
      )
        return;
      setSelection({ slide: slideIndex, index });
      setTarget(element);
      element.tabIndex = -1;
      return candidate;
    },
    [parsed],
  );

  const beginDrag = React.useCallback(
    (
      event: PointerEvent,
      element: HTMLElement,
      candidate: SlideNode,
      corner?: { x: -1 | 0 | 1; y: -1 | 0 | 1 },
    ) => {
      if (busy || event.button !== 0) return;
      if (corner ? candidate.editing?.locks?.noResize : candidate.editing?.locks?.noMove) return;
      const surface = element.closest<HTMLElement>('[data-rpv-slide-index]');
      if (!surface || !parsed) return;
      inline.current?.finish();
      element.focus({ preventScroll: true });
      event.preventDefault();
      const bounds = surface.getBoundingClientRect();
      const scaleX = parsed.document.size.widthEmu / bounds.width;
      const scaleY = parsed.document.size.heightEmu / bounds.height;
      const original = { ...candidate.transform };
      let next: Rect = original;
      const startX = event.clientX;
      const startY = event.clientY;
      const originalStyle = element.getAttribute('style');
      const angle = element.style.transform.includes('rotate')
        ? ((original.rotation ?? 0) * Math.PI) / 180
        : 0;
      let moved = false;
      const move = (current: PointerEvent) => {
        if (current.pointerId !== event.pointerId) return;
        if (!moved && Math.hypot(current.clientX - startX, current.clientY - startY) < 3) return;
        moved = true;
        const dx = (current.clientX - startX) * scaleX;
        const dy = (current.clientY - startY) * scaleY;
        if (corner) {
          const localX = dx * Math.cos(angle) + dy * Math.sin(angle);
          const localY = -dx * Math.sin(angle) + dy * Math.cos(angle);
          let width = Math.max(EMU / 100, original.width + corner.x * localX);
          let height = Math.max(EMU / 100, original.height + corner.y * localY);
          if (
            candidate.editing?.locks?.noChangeAspect ||
            (current.shiftKey && corner.x && corner.y)
          ) {
            const ratio = !corner.x
              ? height / original.height
              : !corner.y
                ? width / original.width
                : Math.max(width / original.width, height / original.height);
            width = original.width * ratio;
            height = original.height * ratio;
          }
          const cx = (corner.x * (width - original.width)) / 2;
          const cy = (corner.y * (height - original.height)) / 2;
          next = {
            x:
              original.x +
              (original.width - width) / 2 +
              cx * Math.cos(angle) -
              cy * Math.sin(angle),
            y:
              original.y +
              (original.height - height) / 2 +
              cx * Math.sin(angle) +
              cy * Math.cos(angle),
            width,
            height,
          };
          if (candidate.editing?.locks?.noMove) {
            next.x = original.x;
            next.y = original.y;
          }
        } else {
          const horizontal = current.shiftKey && Math.abs(dx) >= Math.abs(dy);
          const vertical = current.shiftKey && !horizontal;
          next = {
            ...original,
            x: original.x + (vertical ? 0 : dx),
            y: original.y + (horizontal ? 0 : dy),
          };
        }
        element.style.left = `${(next.x / parsed.document.size.widthEmu) * 100}%`;
        element.style.top = `${(next.y / parsed.document.size.heightEmu) * 100}%`;
        element.style.width = `${(next.width / parsed.document.size.widthEmu) * 100}%`;
        element.style.height = `${(next.height / parsed.document.size.heightEmu) * 100}%`;
        setPreview(next);
      };
      const restore = () => {
        if (element.isConnected && originalStyle !== null)
          element.setAttribute('style', originalStyle);
        setPreview(null);
      };
      const cleanup = (restorePreview = true) => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', finish);
        window.removeEventListener('pointercancel', cancel);
        window.removeEventListener('keydown', escape, true);
        window.removeEventListener('blur', cancel);
        if (restorePreview) restore();
        cancelDrag.current = null;
      };
      const cancel = () => cleanup();
      const finish = (current: PointerEvent) => {
        if (current.pointerId !== event.pointerId) return;
        cleanup(!moved);
        if (moved) void onEdit({ node: candidate, bounds: next }).finally(restore);
      };
      const escape = (current: KeyboardEvent) => {
        if (current.key !== 'Escape') return;
        current.preventDefault();
        current.stopImmediatePropagation();
        cleanup();
      };
      cancelDrag.current = cancel;
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', finish);
      window.addEventListener('pointercancel', cancel);
      window.addEventListener('keydown', escape, true);
      window.addEventListener('blur', cancel);
    },
    [busy, onEdit, parsed],
  );

  const editText = React.useCallback(
    (body: HTMLElement) => {
      if (inline.current?.body === body) return;
      inline.current?.finish();
      const element = body.closest<HTMLElement>('[data-rpv-node-path]');
      const surface = body.closest<HTMLElement>('[data-rpv-slide-index]');
      if (!element || !surface || !parsed) return;
      const path = element.dataset.rpvNodePath?.split('.').map(Number) ?? [];
      let nodes = parsed.document.slides[Number(surface.dataset.rpvSlideIndex)]?.nodes;
      let candidate: SlideNode | undefined;
      for (const index of path) {
        candidate = nodes?.[index];
        if (
          !candidate ||
          candidate.editing?.locks?.noSelection ||
          candidate.editing?.locks?.noTextEdit ||
          candidate.editing?.locks?.noDrilldown
        )
          return;
        nodes = candidate.type === 'group' ? candidate.children : undefined;
      }
      if (
        !candidate ||
        !candidate.sourceId ||
        candidate.sourcePart !==
          parsed.document.slides[Number(surface.dataset.rpvSlideIndex)]?.sourcePart
      )
        return;
      inline.current = startInlineTextEdit(body, candidate, onEdit, () => {
        inline.current = null;
        setTextMode(false);
      });
      setTextMode(true);
      body.focus({ preventScroll: true });
    },
    [parsed, onEdit],
  );

  React.useEffect(() => {
    if (disabled) inline.current?.finish();
  }, [disabled]);

  React.useEffect(() => () => inline.current?.finish(false), []);

  React.useEffect(() => {
    if (disabled) {
      setSelection(null);
      setTarget(null);
    }
    return () => cancelDrag.current?.();
  }, [disabled, parsed]);

  React.useEffect(() => {
    if (slideIndex !== undefined && selection?.slide !== slideIndex) {
      setSelection(null);
      setTarget(null);
    }
  }, [slideIndex, selection?.slide]);

  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || disabled) return;
    const pointerDown = (event: PointerEvent) => {
      if (event.button !== 0 || !(event.target instanceof Element)) return;
      if (event.target.closest('[data-object-controls]')) return;
      if (inline.current?.body.contains(event.target)) return;
      inline.current?.finish();
      const surface = event.target.closest<HTMLElement>('[data-rpv-slide-index]');
      if (!surface) {
        setSelection(null);
        setTarget(null);
        return;
      }
      let element = event.target.closest<HTMLElement>('[data-rpv-node-path]');
      while (element?.parentElement?.closest('[data-rpv-node-path]')) {
        element = element.parentElement.closest<HTMLElement>('[data-rpv-node-path]');
      }
      if (!element) {
        setSelection(null);
        setTarget(null);
        return;
      }
      const candidate = select(element);
      if (!candidate) return;
      const body = event.target.closest<HTMLElement>(
        '[data-rpv-text-content], [data-rpv-table-text-body]',
      );
      if (body && !candidate.editing?.locks?.noTextEdit) {
        editText(body);
        return;
      }
      beginDrag(event, element, candidate);
    };
    const preventNavigation = (event: MouseEvent) => {
      if ((event.target as Element)?.closest('[data-rpv-node-path] a, a[data-rpv-node-path]')) {
        event.preventDefault();
        event.stopPropagation();
      }
    };
    root.addEventListener('pointerdown', pointerDown, true);
    root.addEventListener('click', preventNavigation, true);
    return () => {
      root.removeEventListener('pointerdown', pointerDown, true);
      root.removeEventListener('click', preventNavigation, true);
    };
  }, [rootRef, disabled, busy, select, beginDrag, editText]);

  React.useEffect(() => {
    if (!selection || disabled) return;
    const root = rootRef.current;
    if (!root) return;
    const update = () => {
      const element = root.querySelector<HTMLElement>(
        `[data-rpv-slide-index="${selection.slide}"] [data-rpv-node-path="${selection.index}"]`,
      );
      setTarget(element);
    };
    update();
    const observer = new MutationObserver(update);
    observer.observe(root, { childList: true, subtree: true });
    return () => observer.disconnect();
  }, [selection, parsed, disabled, revision, rootRef]);

  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || disabled || !node || !target) return;
    const keydown = (event: KeyboardEvent) => {
      if (
        !(event.target instanceof HTMLElement) ||
        !root.contains(event.target) ||
        event.target.isContentEditable ||
        event.target instanceof HTMLInputElement ||
        event.target.closest('[data-object-controls]')
      )
        return;
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        setSelection(null);
        setTarget(null);
        return;
      }
      if (event.key === 'Tab') {
        event.preventDefault();
        event.stopPropagation();
        const elements = Array.from(
          target
            .closest('[data-rpv-slide-index]')
            ?.querySelectorAll<HTMLElement>('[data-rpv-node-path]') ?? [],
        ).filter((element) => !element.dataset.rpvNodePath?.includes('.'));
        const current = elements.indexOf(target);
        for (let offset = 1; offset <= elements.length; offset++) {
          const next =
            elements[
              (current + (event.shiftKey ? -offset : offset) + elements.length) % elements.length
            ];
          if (next && select(next)) {
            next.focus({ preventScroll: true });
            break;
          }
        }
        return;
      }
      if (
        event.key === 'Enter' ||
        event.key === 'F2' ||
        (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey)
      ) {
        const body = target.querySelector<HTMLElement>(
          '[data-rpv-text-content], [data-rpv-table-text-body]',
        );
        if (!body) return;
        event.preventDefault();
        event.stopPropagation();
        editText(body);
        if (!body.isContentEditable) return;
        const range = document.createRange();
        range.selectNodeContents(body);
        if (event.key.length !== 1) range.collapse(false);
        window.getSelection()?.removeAllRanges();
        window.getSelection()?.addRange(range);
        if (event.key.length === 1) document.execCommand('insertText', false, event.key);
        return;
      }
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
      event.preventDefault();
      event.stopPropagation();
      if (busy || node.editing?.locks?.noMove) return;
      const delta = ((event.shiftKey ? 10 : 1) * EMU) / 96;
      void onEdit({
        node,
        bounds: {
          ...node.transform,
          x:
            node.transform.x +
            (event.key === 'ArrowLeft' ? -delta : event.key === 'ArrowRight' ? delta : 0),
          y:
            node.transform.y +
            (event.key === 'ArrowUp' ? -delta : event.key === 'ArrowDown' ? delta : 0),
        },
      });
    };
    root.addEventListener('keydown', keydown, true);
    return () => root.removeEventListener('keydown', keydown, true);
  }, [rootRef, disabled, node, target, busy, onEdit, select, editText]);

  if (disabled || !node || !parsed) return null;
  const transform = preview ?? node.transform;
  const objectStyle: React.CSSProperties = {
    position: 'absolute',
    left: `${(transform.x / parsed.document.size.widthEmu) * 100}%`,
    top: `${(transform.y / parsed.document.size.heightEmu) * 100}%`,
    width: preview
      ? `${(transform.width / parsed.document.size.widthEmu) * 100}%`
      : (target?.offsetWidth ?? `${(transform.width / parsed.document.size.widthEmu) * 100}%`),
    height: preview
      ? `${(transform.height / parsed.document.size.heightEmu) * 100}%`
      : (target?.offsetHeight ?? `${(transform.height / parsed.document.size.heightEmu) * 100}%`),
    transform: target?.style.transform.replace(/scale[XY]\(-1\)/g, ''),
    border: `1px ${textMode ? 'dashed' : 'solid'} #737373`,
    pointerEvents: 'none',
    zIndex: 10000,
  };
  return (
    <>
      {target &&
        slide &&
        createPortal(
          <div data-object-controls style={objectStyle}>
            {(['top', 'right', 'bottom', 'left'] as const).map((edge) => (
              <div
                key={edge}
                role="button"
                aria-label={`Move object from ${edge} border`}
                style={{
                  position: 'absolute',
                  pointerEvents: 'auto',
                  cursor: 'move',
                  touchAction: 'none',
                  ...(edge === 'top' || edge === 'bottom'
                    ? { left: 0, right: 0, height: 8, [edge]: -4 }
                    : { top: 0, bottom: 0, width: 8, [edge]: -4 }),
                }}
                onPointerDown={(event) => {
                  event.stopPropagation();
                  beginDrag(event.nativeEvent, target, node);
                }}
              />
            ))}
            {!node.editing?.locks?.noResize &&
              (
                [
                  [-1, -1],
                  [0, -1],
                  [1, -1],
                  [-1, 0],
                  [1, 0],
                  [-1, 1],
                  [0, 1],
                  [1, 1],
                ] as const
              ).map(([x, y]) => (
                <button
                  key={`${x},${y}`}
                  type="button"
                  aria-label={`Resize ${y < 0 ? 'top' : y > 0 ? 'bottom' : ''} ${x < 0 ? 'left' : x > 0 ? 'right' : ''}`.trim()}
                  style={{
                    position: 'absolute',
                    left: x < 0 ? 0 : x > 0 ? '100%' : '50%',
                    top: y < 0 ? 0 : y > 0 ? '100%' : '50%',
                    transform: 'translate(-50%, -50%)',
                    width: 8,
                    height: 8,
                    padding: 0,
                    background: 'white',
                    border: '1px solid #737373',
                    pointerEvents: 'auto',
                    touchAction: 'none',
                    cursor: !x
                      ? 'ns-resize'
                      : !y
                        ? 'ew-resize'
                        : x === y
                          ? 'nwse-resize'
                          : 'nesw-resize',
                  }}
                  onPointerDown={(event) => {
                    event.stopPropagation();
                    beginDrag(event.nativeEvent, target, node, { x, y });
                  }}
                  onKeyDown={(event) => {
                    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key))
                      return;
                    event.preventDefault();
                    event.stopPropagation();
                    const delta = ((event.shiftKey ? 10 : 1) * EMU) / 96;
                    const width = Math.max(
                      EMU / 100,
                      node.transform.width +
                        (event.key === 'ArrowRight'
                          ? delta
                          : event.key === 'ArrowLeft'
                            ? -delta
                            : 0),
                    );
                    const height = node.editing?.locks?.noChangeAspect
                      ? (width * node.transform.height) / node.transform.width
                      : Math.max(
                          EMU / 100,
                          node.transform.height +
                            (event.key === 'ArrowDown'
                              ? delta
                              : event.key === 'ArrowUp'
                                ? -delta
                                : 0),
                        );
                    void onEdit({ node, bounds: { ...node.transform, width, height } });
                  }}
                />
              ))}
          </div>,
          slide,
        )}
    </>
  );
}
