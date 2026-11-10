import { useEffect, useMemo, useRef, useState, type RefObject } from 'react';
import { useVirtualizer } from '@tanstack/react-virtual';
import {
  extractPresentationBlocks,
  PptxBlockHighlights,
  revealPptxSource,
  type PresentationDocument,
  type PresentationBlock,
  type PresentationWord,
  type StructuredPresentation,
  type PptxViewerController,
  type PptxBlockGeometry,
} from '@extend-ai/react-pptx';
import { Download, ScanText, Search, X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { InputGroup, InputGroupAddon, InputGroupInput } from '@/components/ui/input-group';
import { Label } from '@/components/ui/label';
import { SegmentedControl, SegmentedControlItem } from '@/components/ui/segmented-control';
import {
  Select,
  SelectItem,
  SelectPopup,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { Spinner } from '@/components/ui/spinner';
import { Switch } from '@/components/ui/switch';
import { Tooltip, TooltipPopup, TooltipTrigger } from '@/components/ui/tooltip';
import { cn } from '@/lib/utils';

type Row = { target: PresentationBlock | PresentationWord; type: string; depth: number };
const colors: Record<string, string> = {
  text: '#2563eb',
  text_box: '#2563eb',
  table: '#047857',
  table_cell: '#047857',
  figure: '#a16207',
  group: '#9333ea',
  note: '#64748b',
  word: '#2563eb',
};
function rows(blocks: readonly PresentationBlock[], depth = 0): Row[] {
  return blocks.flatMap((block) => [
    { target: block, type: block.type, depth },
    ...rows(block.children, depth + 1),
  ]);
}

export function BlockInspector({
  model,
  rootRef,
  controller,
  onClose,
}: {
  model: PresentationDocument;
  rootRef: RefObject<HTMLDivElement | null>;
  controller: PptxViewerController | null;
  onClose: () => void;
}) {
  const [output, setOutput] = useState<StructuredPresentation>();
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(true);
  const [mode, setMode] = useState<'text' | 'words' | 'json'>('text');
  const [includeWords, setIncludeWords] = useState(false);
  const [type, setType] = useState('all');
  const [query, setQuery] = useState('');
  const [showAll, setShowAll] = useState(false);
  const [hovered, setHovered] = useState<string>();
  const [selected, setSelected] = useState<string>();
  const [geometry, setGeometry] = useState<PptxBlockGeometry>();
  const scrollRef = useRef<HTMLDivElement>(null);
  const navigation = useRef(0);
  useEffect(() => {
    let cancelled = false;
    setBusy(true);
    setOutput(undefined);
    setError(undefined);
    setHovered(undefined);
    setSelected(undefined);
    setGeometry(undefined);
    navigation.current += 1;
    void extractPresentationBlocks(model, { includeWords, includeNotes: true })
      .then((result) => {
        if (!cancelled) {
          setOutput(result);
          setBusy(false);
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setError(String(error));
          setBusy(false);
        }
      });
    return () => {
      cancelled = true;
      navigation.current += 1;
    };
  }, [model, includeWords]);
  const all = useMemo(() => rows(output?.blocks ?? []), [output]);
  const types = useMemo(() => [...new Set(all.map((row) => row.type))], [all]);
  const filtered = useMemo(() => {
    const blocks = all.filter((row) => type === 'all' || row.type === type);
    const candidates =
      mode === 'words'
        ? blocks.flatMap((row) =>
            row.target.object === 'block'
              ? (row.target.words ?? []).map((word) => ({
                  target: word,
                  type: 'word',
                  depth: row.depth,
                }))
              : [],
          )
        : blocks;
    return candidates.filter((row) =>
      row.target.content.toLocaleLowerCase().includes(query.toLocaleLowerCase()),
    );
  }, [all, type, mode, query]);
  const active =
    mode === 'json'
      ? all
          .flatMap((row) => [
            row.target,
            ...(row.target.object === 'block' ? (row.target.words ?? []) : []),
          ])
          .find((target) => target.id === selected)
      : filtered.find((row) => row.target.id === (hovered ?? selected))?.target;
  const highlights = useMemo(
    () => (showAll ? filtered.map((row) => row.target) : []),
    [showAll, filtered],
  );
  const colorById = useMemo(
    () => new Map(filtered.map((row) => [row.target.id, colors[row.type] ?? '#2563eb'])),
    [filtered],
  );
  const virtualizer = useVirtualizer({
    count: filtered.length,
    getItemKey: (i) => filtered[i]!.target.id,
    getScrollElement: () => scrollRef.current,
    estimateSize: () => (mode === 'words' ? 58 : 104),
    overscan: 5,
  });
  useEffect(() => {
    virtualizer.measure();
  }, [filtered, mode, virtualizer]);
  const select = async (target: PresentationBlock | PresentationWord) => {
    setSelected(target.id);
    setGeometry(undefined);
    const generation = ++navigation.current;
    const source = target.sourceRefs[0];
    if (!source || !controller || !rootRef.current) return;
    try {
      const result = await revealPptxSource(controller, rootRef.current, source);
      if (generation === navigation.current) setGeometry(result);
    } catch (error) {
      if (generation === navigation.current) setError(String(error));
    }
  };
  const download = () => {
    if (!output) return;
    const url = URL.createObjectURL(
      new Blob([JSON.stringify(output, null, 2)], { type: 'application/json' }),
    );
    const link = document.createElement('a');
    link.href = url;
    link.download = 'presentation-blocks.json';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const status =
    !active || !geometry
      ? undefined
      : geometry.status === 'rendered'
        ? `Slide ${active.sourceRefs[0]!.slideIndex + 1} · ${geometry.fragments.length} rectangles`
        : geometry.issues.some((issue) => issue.reason === 'not_on_slide')
          ? 'Speaker notes are not drawn on the slide.'
          : geometry.status === 'not_mounted'
            ? 'Click to reveal this slide.'
            : geometry.status === 'not_visible'
              ? 'This content is not visible on the slide.'
              : 'This source has no supported slide geometry.';
  return (
    <>
      <aside
        aria-label="Parsed blocks"
        className="flex w-[340px] shrink-0 flex-col overflow-hidden rounded-lg border bg-background text-[13px] shadow-xs/5"
      >
        <header className="flex h-10 shrink-0 items-center gap-2 border-b ps-3 pe-1.5">
          <ScanText className="size-4 text-brand" />
          <strong className="font-semibold">Blocks</strong>
          <span aria-live="polite" className="text-xs text-muted-foreground">
            {busy ? 'Loading…' : `${filtered.length} ${mode === 'words' ? 'words' : 'blocks'}`}
          </span>
          <span className="flex-1" />
          <Tooltip>
            <TooltipTrigger
              render={
                <Button
                  aria-label="Download JSON"
                  className="size-7 sm:size-7"
                  disabled={!output || busy}
                  onClick={download}
                  size="icon-sm"
                  variant="ghost"
                />
              }
            >
              <Download />
            </TooltipTrigger>
            <TooltipPopup side="bottom">Download JSON</TooltipPopup>
          </Tooltip>
          <Button
            aria-label="Close blocks"
            className="size-7 sm:size-7"
            onClick={onClose}
            size="icon-sm"
            variant="ghost"
          >
            <X />
          </Button>
        </header>
        <div className="flex shrink-0 flex-col gap-2 border-b p-3">
          <div className="flex items-center gap-2">
            <SegmentedControl aria-label="Block view" className="flex-1" size="sm">
              {(['text', 'words', 'json'] as const).map((value) => (
                <SegmentedControlItem
                  aria-pressed={mode === value}
                  className="flex-1 text-[13px] sm:text-[13px]"
                  data-pressed={mode === value ? '' : undefined}
                  key={value}
                  onClick={() => {
                    setMode(value);
                    setHovered(undefined);
                    if (value !== 'json') setSelected(undefined);
                    setGeometry(undefined);
                    if (value === 'words') setIncludeWords(true);
                  }}
                >
                  {value === 'text' ? 'Text' : value === 'words' ? 'Words' : 'JSON'}
                </SegmentedControlItem>
              ))}
            </SegmentedControl>
            <Select
              items={[
                { label: 'All types', value: 'all' },
                ...types.map((value) => ({ label: value.replaceAll('_', ' '), value })),
              ]}
              onValueChange={(value) => {
                setType(value ?? 'all');
                setHovered(undefined);
                setSelected(undefined);
              }}
              value={type}
            >
              <SelectTrigger
                aria-label="Filter block type"
                className="w-32 min-w-0 text-[13px] capitalize sm:text-[13px]"
                size="sm"
              >
                <SelectValue />
              </SelectTrigger>
              <SelectPopup>
                <SelectItem value="all">All types</SelectItem>
                {types.map((value) => (
                  <SelectItem className="capitalize" key={value} value={value}>
                    {value.replaceAll('_', ' ')}
                  </SelectItem>
                ))}
              </SelectPopup>
            </Select>
          </div>
          <InputGroup className="rounded-md before:rounded-[calc(var(--radius-md)-1px)]">
            <InputGroupAddon>
              <Search className="size-3.5!" />
            </InputGroupAddon>
            <InputGroupInput
              aria-label="Search blocks"
              className="text-[13px] sm:text-[13px]"
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Filter block text"
              size="sm"
              value={query}
            />
          </InputGroup>
          <Label className="cursor-pointer gap-2 text-[13px] font-normal">
            <Switch
              checked={showAll}
              className="data-checked:bg-brand"
              onCheckedChange={setShowAll}
            />
            Show all highlights
          </Label>
        </div>
        {error ? (
          <p className="border-b px-3 py-2 text-xs text-destructive-foreground" role="alert">
            {error}
          </p>
        ) : null}
        <div
          className="min-h-0 flex-1 overflow-auto"
          ref={scrollRef}
          onMouseLeave={() => setHovered(undefined)}
        >
          {busy ? (
            <p className="flex items-center gap-2 p-3 text-xs text-muted-foreground">
              <Spinner className="size-3.5" />
              Extracting blocks…
            </p>
          ) : mode === 'json' ? (
            <pre className="m-0 p-3 font-mono text-[11px] leading-relaxed whitespace-pre-wrap break-all">
              {JSON.stringify(active ? { ...active, geometry } : output, null, 2)}
            </pre>
          ) : filtered.length ? (
            <div style={{ height: virtualizer.getTotalSize(), position: 'relative' }}>
              {virtualizer.getVirtualItems().map((item) => {
                const row = filtered[item.index]!,
                  target = row.target,
                  source = target.sourceRefs[0];
                const isActive = active?.id === target.id;
                return (
                  <button
                    key={item.key}
                    type="button"
                    ref={virtualizer.measureElement}
                    data-index={item.index}
                    data-block-output-id={target.id}
                    data-block-output-type={row.type}
                    data-active={isActive || undefined}
                    className={cn(
                      'absolute inset-x-0 top-0 flex w-full flex-col gap-1 border-b border-s-2 border-s-transparent py-2.5 pe-3 text-left outline-none transition-colors hover:bg-accent focus-visible:bg-accent',
                      isActive && 'border-s-brand bg-brand/6 hover:bg-brand/8',
                    )}
                    style={{
                      transform: `translateY(${item.start}px)`,
                      paddingLeft: 12 + Math.min(row.depth, 4) * 12,
                    }}
                    onMouseEnter={() => {
                      setHovered(target.id);
                      setGeometry(undefined);
                    }}
                    onFocus={() => {
                      setHovered(target.id);
                      setGeometry(undefined);
                    }}
                    onBlur={() => setHovered(undefined)}
                    onClick={() => void select(target)}
                  >
                    <span className="flex items-center gap-1.5 text-[11px]">
                      <span
                        className="size-2 rounded-full"
                        style={{ backgroundColor: colors[row.type] ?? '#2563eb' }}
                      />
                      <span className="font-medium capitalize" style={{ color: colors[row.type] }}>
                        {row.type.replaceAll('_', ' ')}
                      </span>
                      <span className="ms-auto text-muted-foreground">
                        {source?.story === 'notes' ? 'Notes · ' : ''}Slide{' '}
                        {(source?.slideIndex ?? 0) + 1}
                      </span>
                    </span>
                    <span className="line-clamp-3 text-[13px] leading-snug break-words">
                      {target.content ||
                        (target.object === 'block' ? (
                          <span className="text-muted-foreground capitalize">
                            {String(target.details.kind ?? target.type).replaceAll('_', ' ')}
                          </span>
                        ) : (
                          <span className="text-muted-foreground">Empty word</span>
                        ))}
                    </span>
                  </button>
                );
              })}
            </div>
          ) : (
            <p className="p-3 text-xs text-muted-foreground">No matching blocks.</p>
          )}
        </div>
        <footer
          aria-live="polite"
          className="shrink-0 border-t bg-chrome/60 px-3 py-2 text-[11px] text-muted-foreground"
        >
          {status ?? 'Hover to highlight. Click to reveal on the slide.'}
        </footer>
      </aside>
      <PptxBlockHighlights
        rootRef={rootRef}
        {...(active ? { block: active } : {})}
        blocks={highlights}
        colorForBlock={(target) => colorById.get(target.id) ?? '#2563eb'}
        onGeometry={setGeometry}
      />
    </>
  );
}
