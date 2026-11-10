import * as React from 'react';
import { createPortal } from 'react-dom';
import type { ParsedPresentation } from '@extend-ai/react-pptx';
import type { Rect, SlideNode } from '@extend-ai/react-pptx-model';
import { X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { Label } from '@/components/ui/label';
import type { ObjectEdit } from '@/lib/pptx-editing';

const EMU = 914400;
type Selection = { slide: number; index: number };
type Props = {
  rootRef: React.RefObject<HTMLDivElement | null>;
  parsed: ParsedPresentation | null;
  disabled: boolean;
  busy: boolean;
  revision: number;
  onEdit: (edit: ObjectEdit) => Promise<void>;
};

function textNodes(node: SlideNode): SlideNode[] {
  if (
    node.editing?.locks?.noTextEdit ||
    node.editing?.locks?.noSelection ||
    node.editing?.locks?.noDrilldown
  )
    return [];
  if (node.type === 'group') return node.children.flatMap(textNodes);
  return node.type === 'shape' || node.type === 'table' ? [node] : [];
}

function ParagraphField({
  node,
  index,
  text,
  row,
  column,
  busy,
  onEdit,
}: {
  node: SlideNode;
  index: number;
  text: string;
  row?: number;
  column?: number;
  busy: boolean;
  onEdit: Props['onEdit'];
}) {
  const id = React.useId();
  const [value, setValue] = React.useState(text);
  React.useEffect(() => setValue(text), [text]);
  return (
    <form
      className="space-y-2"
      onSubmit={(event) => {
        event.preventDefault();
        void onEdit({ node, paragraph: { index, text: value, row, column } });
      }}
    >
      <Label className="text-xs" htmlFor={id}>
        {row === undefined || column === undefined
          ? `Text ${index + 1}`
          : `Row ${row + 1}, column ${column + 1}, text ${index + 1}`}
      </Label>
      <Textarea
        id={id}
        className="min-h-20 text-xs"
        value={value}
        disabled={busy}
        onChange={(event) => setValue(event.target.value)}
      />
      <Button type="submit" size="sm" variant="outline" disabled={busy || value === text}>
        Apply text
      </Button>
    </form>
  );
}

function BoundsFields({
  node,
  busy,
  onEdit,
}: {
  node: SlideNode;
  busy: boolean;
  onEdit: Props['onEdit'];
}) {
  const id = React.useId();
  const [values, setValues] = React.useState(() => ({ ...node.transform }));
  React.useEffect(() => setValues({ ...node.transform }), [node]);
  const locks = node.editing?.locks;
  const changed = (['x', 'y', 'width', 'height'] as const).some(
    (key) => values[key] !== node.transform[key],
  );
  return (
    <form
      className="space-y-2"
      onSubmit={(event) => {
        event.preventDefault();
        void onEdit({ node, bounds: values });
      }}
    >
      <div className="grid grid-cols-2 gap-2">
        {(['x', 'y', 'width', 'height'] as const).map((key) => (
          <div key={key} className="space-y-1">
            <Label className="text-xs capitalize" htmlFor={`${id}-${key}`}>
              {key} (in)
            </Label>
            <Input
              id={`${id}-${key}`}
              type="number"
              step="0.001"
              className="h-8 text-xs"
              min={key === 'width' || key === 'height' ? 0.01 : undefined}
              disabled={busy || (key === 'x' || key === 'y' ? locks?.noMove : locks?.noResize)}
              value={Number((values[key] / EMU).toFixed(3))}
              onChange={(event) => {
                const value = Number(event.target.value) * EMU;
                setValues((current) => {
                  const next = { ...current, [key]: value };
                  if (locks?.noChangeAspect && key === 'width')
                    next.height = (value * node.transform.height) / node.transform.width;
                  if (locks?.noChangeAspect && key === 'height')
                    next.width = (value * node.transform.width) / node.transform.height;
                  return next;
                });
              }}
            />
          </div>
        ))}
      </div>
      <Button type="submit" size="sm" variant="outline" disabled={busy || !changed}>
        Apply position and size
      </Button>
    </form>
  );
}

export function ObjectEditor({ rootRef, parsed, disabled, busy, revision, onEdit }: Props) {
  const [selection, setSelection] = React.useState<Selection | null>(null);
  const [target, setTarget] = React.useState<HTMLElement | null>(null);
  const [preview, setPreview] = React.useState<Rect | null>(null);
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
      return candidate;
    },
    [parsed],
  );

  const beginDrag = React.useCallback(
    (
      event: PointerEvent,
      element: HTMLElement,
      candidate: SlideNode,
      corner?: { x: -1 | 1; y: -1 | 1 },
    ) => {
      if (busy || event.button !== 0) return;
      if (corner ? candidate.editing?.locks?.noResize : candidate.editing?.locks?.noMove) return;
      const surface = element.closest<HTMLElement>('[data-rpv-slide-index]');
      if (!surface || !parsed) return;
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
          if (candidate.editing?.locks?.noChangeAspect || current.shiftKey) {
            const ratio = Math.max(width / original.width, height / original.height);
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
          next = { ...original, x: original.x + dx, y: original.y + dy };
        }
        element.style.left = `${(next.x / parsed.document.size.widthEmu) * 100}%`;
        element.style.top = `${(next.y / parsed.document.size.heightEmu) * 100}%`;
        element.style.width = `${(next.width / parsed.document.size.widthEmu) * 100}%`;
        element.style.height = `${(next.height / parsed.document.size.heightEmu) * 100}%`;
        setPreview(next);
      };
      const cleanup = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', finish);
        window.removeEventListener('pointercancel', cancel);
        window.removeEventListener('keydown', escape, true);
        window.removeEventListener('blur', cancel);
        if (originalStyle !== null) element.setAttribute('style', originalStyle);
        setPreview(null);
        cancelDrag.current = null;
      };
      const cancel = () => cleanup();
      const finish = (current: PointerEvent) => {
        if (current.pointerId !== event.pointerId) return;
        cleanup();
        if (moved) void onEdit({ node: candidate, bounds: next });
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

  React.useEffect(() => {
    if (disabled) {
      setSelection(null);
      setTarget(null);
    }
    return () => cancelDrag.current?.();
  }, [disabled, parsed]);

  React.useEffect(() => {
    const root = rootRef.current;
    if (!root || disabled) return;
    const pointerDown = (event: PointerEvent) => {
      if (event.button !== 0 || !(event.target instanceof Element)) return;
      if (event.target.closest('[data-object-controls]')) return;
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
      if (candidate) beginDrag(event, element, candidate);
    };
    const preventNavigation = (event: MouseEvent) => {
      if ((event.target as Element)?.closest('[data-rpv-node-path]')) {
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
  }, [rootRef, disabled, select, beginDrag]);

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

  if (disabled || !node || !target || !slide || !parsed) return null;
  const transform = preview ?? node.transform;
  const objectStyle: React.CSSProperties = {
    position: 'absolute',
    left: `${(transform.x / parsed.document.size.widthEmu) * 100}%`,
    top: `${(transform.y / parsed.document.size.heightEmu) * 100}%`,
    width: preview
      ? `${(transform.width / parsed.document.size.widthEmu) * 100}%`
      : target.offsetWidth,
    height: preview
      ? `${(transform.height / parsed.document.size.heightEmu) * 100}%`
      : target.offsetHeight,
    transform: target.style.transform.replace(/scale[XY]\(-1\)/g, ''),
    border: '1.5px solid #2563eb',
    pointerEvents: 'none',
    zIndex: 10000,
  };
  return (
    <>
      {createPortal(
        <div data-object-controls style={objectStyle}>
          {!node.editing?.locks?.noResize &&
            !busy &&
            (
              [
                [-1, -1],
                [1, -1],
                [-1, 1],
                [1, 1],
              ] as const
            ).map(([x, y]) => (
              <button
                key={`${x},${y}`}
                type="button"
                aria-label={`Resize ${y < 0 ? 'top' : 'bottom'} ${x < 0 ? 'left' : 'right'}`}
                style={{
                  position: 'absolute',
                  left: x < 0 ? 0 : '100%',
                  top: y < 0 ? 0 : '100%',
                  transform: 'translate(-50%, -50%)',
                  width: 10,
                  height: 10,
                  padding: 0,
                  background: 'white',
                  border: '1.5px solid #2563eb',
                  pointerEvents: 'auto',
                  touchAction: 'none',
                  cursor: x === y ? 'nwse-resize' : 'nesw-resize',
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
                      (event.key === 'ArrowRight' ? delta : event.key === 'ArrowLeft' ? -delta : 0),
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
      <aside
        aria-label="Edit object"
        className="flex w-64 shrink-0 flex-col overflow-hidden rounded-lg border bg-background max-md:absolute max-md:right-3 max-md:top-32 max-md:bottom-10 max-md:z-20 max-md:shadow-lg"
      >
        <div className="flex items-center justify-between border-b p-3">
          <span className="text-sm font-medium">Edit object</span>
          <Button
            size="icon-sm"
            variant="ghost"
            aria-label="Close object editor"
            onClick={() => {
              setSelection(null);
              setTarget(null);
            }}
          >
            <X />
          </Button>
        </div>
        <div className="space-y-5 overflow-y-auto p-3">
          <p className="text-xs text-muted-foreground">
            Drag to move. Use the corner handles to resize. Hold Shift to keep proportions.
          </p>
          <BoundsFields node={node} busy={busy} onEdit={onEdit} />
          {textNodes(node).flatMap((entry) =>
            entry.type === 'shape'
              ? (entry.paragraphs ?? []).map((paragraph, index) => (
                  <ParagraphField
                    key={`${entry.id}:${index}`}
                    node={entry}
                    index={index}
                    text={paragraph.runs.map((run) => run.text).join('')}
                    busy={busy}
                    onEdit={onEdit}
                  />
                ))
              : entry.type === 'table'
                ? entry.rows.flatMap((row, rowIndex) =>
                    row.flatMap((cell, column) =>
                      cell.merge?.horizontal || cell.merge?.vertical
                        ? []
                        : cell.paragraphs.map((paragraph, index) => (
                            <ParagraphField
                              key={`${entry.id}:${rowIndex}:${column}:${index}`}
                              node={entry}
                              index={index}
                              row={rowIndex}
                              column={column}
                              text={paragraph.runs.map((run) => run.text).join('')}
                              busy={busy}
                              onEdit={onEdit}
                            />
                          )),
                    ),
                  )
                : [],
          )}
        </div>
      </aside>
    </>
  );
}
