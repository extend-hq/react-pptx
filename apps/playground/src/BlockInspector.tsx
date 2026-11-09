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
      <aside className="blocks-panel" aria-label="Parsed blocks">
        <header className="blocks-panel__header">
          <strong>Blocks</strong>
          <span aria-live="polite">
            {busy ? 'Loading…' : `${filtered.length} ${mode === 'words' ? 'words' : 'blocks'}`}
          </span>
          <button className="button outline small" onClick={download} disabled={!output || busy}>
            Download JSON
          </button>
          <button className="button outline small" onClick={onClose} aria-label="Close blocks">
            ×
          </button>
        </header>
        <div className="blocks-panel__controls">
          <div className="button-group" role="group" aria-label="Block view">
            {(['text', 'words', 'json'] as const).map((value) => (
              <button
                className="button outline small"
                key={value}
                aria-pressed={mode === value}
                data-active={mode === value || undefined}
                onClick={() => {
                  setMode(value);
                  setHovered(undefined);
                  if (value !== 'json') setSelected(undefined);
                  setGeometry(undefined);
                  if (value === 'words') setIncludeWords(true);
                }}
              >
                {value === 'text' ? 'Text' : value === 'words' ? 'Words' : 'JSON'}
              </button>
            ))}
          </div>
          <label>
            Type
            <select
              aria-label="Filter block type"
              value={type}
              onChange={(event) => {
                setType(event.target.value);
                setHovered(undefined);
                setSelected(undefined);
              }}
            >
              <option value="all">All types</option>
              {types.map((value) => (
                <option key={value} value={value}>
                  {value.replaceAll('_', ' ')}
                </option>
              ))}
            </select>
          </label>
          <input
            aria-label="Search blocks"
            placeholder="Filter block text"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
          <label className="blocks-panel__check">
            <input
              type="checkbox"
              checked={showAll}
              onChange={(event) => setShowAll(event.target.checked)}
            />
            Show all highlights
          </label>
        </div>
        {error ? (
          <p role="alert" className="blocks-panel__message">
            {error}
          </p>
        ) : null}
        <div
          className="blocks-panel__list"
          ref={scrollRef}
          onMouseLeave={() => setHovered(undefined)}
        >
          {busy ? (
            <p className="blocks-panel__message">Extracting blocks…</p>
          ) : mode === 'json' ? (
            <pre>{JSON.stringify(active ? { ...active, geometry } : output, null, 2)}</pre>
          ) : filtered.length ? (
            <div style={{ height: virtualizer.getTotalSize(), position: 'relative' }}>
              {virtualizer.getVirtualItems().map((item) => {
                const row = filtered[item.index]!,
                  target = row.target,
                  source = target.sourceRefs[0];
                return (
                  <button
                    key={item.key}
                    type="button"
                    ref={virtualizer.measureElement}
                    data-index={item.index}
                    data-block-output-id={target.id}
                    data-block-output-type={row.type}
                    data-active={active?.id === target.id || undefined}
                    className="block-card"
                    style={{
                      transform: `translateY(${item.start}px)`,
                      paddingLeft: 14 + Math.min(row.depth, 4) * 12,
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
                    <span className="block-card__meta">
                      <span style={{ color: colors[row.type] }}>
                        {row.type.replaceAll('_', ' ')}
                      </span>
                      <span>
                        {source?.story === 'notes' ? 'Notes · ' : ''}Slide{' '}
                        {(source?.slideIndex ?? 0) + 1}
                      </span>
                    </span>
                    <span className="block-card__content">
                      {target.content ||
                        (target.object === 'block'
                          ? String(target.details.kind ?? target.type).replaceAll('_', ' ')
                          : 'Empty word')}
                    </span>
                  </button>
                );
              })}
            </div>
          ) : (
            <p className="blocks-panel__message">No matching blocks.</p>
          )}
        </div>
        <footer className="blocks-panel__message" aria-live="polite">
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
