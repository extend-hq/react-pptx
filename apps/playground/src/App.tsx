import * as React from 'react';
import wasmUrl from '@extend-ai/react-pptx/pptx_wasm_bg.wasm?url';
import { useVirtualizer } from '@tanstack/react-virtual';
import {
  ReactPptxViewer,
  setWasmSource,
  type ParsedPresentation,
  type PresentationSearchResult,
  type PresentationSource,
  type PptxViewerController,
  type ViewerMode,
  type ViewerZoomLevel,
  usePptxViewer,
  usePptxViewerThumbnails,
} from '@extend-ai/react-pptx';
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Download,
  FileUp,
  FolderOpen,
  GalleryVertical,
  Minus,
  Moon,
  MonitorPlay,
  PanelLeft,
  Play,
  Plus,
  Presentation,
  RectangleHorizontal,
  RotateCcw,
  Undo2,
  Redo2,
  Scan,
  ScanText,
  Search,
  StickyNote,
  Sun,
  Upload,
} from 'lucide-react';
import { useTheme } from 'next-themes';
import { BlockInspector } from '@/BlockInspector';
import { ObjectEditor } from '@/ObjectEditor';
import { usePptxEditing } from '@/lib/use-pptx-editing';
import {
  RibbonButton,
  RibbonDivider,
  RibbonGroup,
  RibbonMenuButton,
  RibbonToggle,
} from '@/components/ribbon';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { InputGroup, InputGroupAddon, InputGroupInput } from '@/components/ui/input-group';
import {
  Menu,
  MenuGroup,
  MenuGroupLabel,
  MenuItem,
  MenuPopup,
  MenuSeparator,
  MenuTrigger,
} from '@/components/ui/menu';
import { Slider } from '@/components/ui/slider';
import { Spinner } from '@/components/ui/spinner';
import { Tabs, TabsList, TabsPanel, TabsTab } from '@/components/ui/tabs';
import { Tooltip, TooltipPopup, TooltipTrigger } from '@/components/ui/tooltip';
import { cn } from '@/lib/utils';

setWasmSource(wasmUrl);

const SAMPLES = [
  {
    file: 'generative-ai.pptx',
    label: 'Introduction to Generative AI',
    source: '/samples/generative-ai.pptx',
  },
  { file: 'viewer-smoke.pptx', label: 'Viewer smoke test', source: '/viewer-smoke.pptx' },
  { file: 'charts-showcase.pptx', label: 'Charts showcase', source: '/charts-showcase.pptx' },
] as const;
const DEFAULT_SAMPLE = SAMPLES[0];
const ZOOM_PRESETS = [50, 75, 100, 125, 150, 200, 300] as const;
const MIN_ZOOM = 10;
const MAX_ZOOM = 400;
const ZOOM_STEP = 10;
const THUMBNAIL_WIDTH = 168;
const THUMBNAIL_ROW_ESTIMATE = 112;
const THUMBNAIL_OVERSCAN = 1;
const THUMBNAIL_PREFETCH = 2;
const EMU_PER_INCH = 914_400;
const VIRTUALIZATION = { enabled: true, initialSlides: 3, batchSize: 3, overscanViewport: 1.5 };
const RIBBON_TABS = ['Home', 'Insert', 'Design', 'Slide Show', 'View', 'Developer'] as const;
type RibbonTab = (typeof RIBBON_TABS)[number];

function buildThumbnailRenderWindow(visibleIndexes: readonly number[], slideCount: number) {
  const visibleSlideIndexes = [...new Set(visibleIndexes)]
    .filter((index) => index >= 0 && index < slideCount)
    .sort((left, right) => left - right);
  if (!visibleSlideIndexes.length) {
    return { visibleSlideIndexes, prefetchSlideIndexes: [] };
  }

  const visible = new Set(visibleSlideIndexes);
  const prefetchSlideIndexes: number[] = [];
  const first = visibleSlideIndexes[0] ?? 0;
  const last = visibleSlideIndexes.at(-1) ?? first;
  for (let distance = 1; distance <= THUMBNAIL_PREFETCH; distance += 1) {
    for (const index of [first - distance, last + distance]) {
      if (index >= 0 && index < slideCount && !visible.has(index)) {
        prefetchSlideIndexes.push(index);
      }
    }
  }
  return { visibleSlideIndexes, prefetchSlideIndexes };
}

const clampZoom = (value: number) => Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(value)));

/* -------------------------------------------------------------------------- */
/*                               Playground state                             */
/* -------------------------------------------------------------------------- */

type Playground = {
  controller: PptxViewerController | null;
  parsed: ParsedPresentation | null;
  fileName: string;
  slide: number;
  slideCount: number;
  mode: ViewerMode;
  setMode: (mode: ViewerMode) => void;
  zoom: ViewerZoomLevel;
  setZoom: (zoom: ViewerZoomLevel) => void;
  resolvedZoom: number;
  search: string;
  setSearch: (value: string) => void;
  results: readonly PresentationSearchResult[];
  activeResult: number | null;
  stepResult: (direction: 1 | -1) => void;
  diagnostics: boolean;
  setDiagnostics: (value: boolean) => void;
  thumbnailsOpen: boolean;
  setThumbnailsOpen: (value: boolean) => void;
  blocksOpen: boolean;
  setBlocksOpen: (value: boolean) => void;
  notesOpen: boolean;
  setNotesOpen: (value: boolean) => void;
  slideLabels: boolean;
  setSlideLabels: (value: boolean) => void;
  openFilePicker: () => void;
  loadSample: (file: string) => void;
  resetSample: () => void;
  downloadOriginal: () => void;
  readOnly: boolean;
  setReadOnly: (value: boolean) => void;
  legacy: boolean;
  editingBusy: boolean;
  editingReady: boolean;
  canUndo: boolean;
  canRedo: boolean;
  undo: () => void;
  redo: () => void;
  startSlideShow: (fromBeginning: boolean) => void;
};

const PlaygroundContext = React.createContext<Playground | null>(null);

function usePlayground() {
  const value = React.useContext(PlaygroundContext);
  if (!value) throw new Error('usePlayground must be used inside the playground.');
  return value;
}

/* -------------------------------------------------------------------------- */
/*                                  Title bar                                 */
/* -------------------------------------------------------------------------- */

function TitleBar() {
  const {
    downloadOriginal,
    fileName,
    openFilePicker,
    parsed,
    readOnly,
    editingReady,
  } = usePlayground();
  const { resolvedTheme, setTheme } = useTheme();
  const isDark = resolvedTheme === 'dark';
  const titleButton = 'h-7 text-[13px] font-normal sm:h-7';

  return (
    <header className="flex h-11 shrink-0 items-center gap-3 px-3">
      <div className="flex min-w-0 flex-1 items-center gap-2.5">
        <div className="flex size-7 shrink-0 items-center justify-center rounded-md bg-brand text-brand-foreground">
          <Presentation className="size-4" />
        </div>
        <span className="truncate text-[13px] font-semibold">{fileName}</span>
        {readOnly && (
          <Badge className="rounded px-1.5 text-[10px]" variant="outline">
            Read-only
          </Badge>
        )}
        <span className="hidden truncate text-xs text-muted-foreground md:inline">
          React PPTX Playground
        </span>
      </div>

      <div className="flex shrink-0 items-center gap-0.5">
        <Button className={titleButton} onClick={openFilePicker} size="sm" variant="ghost">
          <FolderOpen />
          Open
        </Button>
        <Button
          className={titleButton}
          disabled={!parsed || !editingReady}
          onClick={downloadOriginal}
          size="sm"
          variant="ghost"
        >
          <Download />
          Download
        </Button>
        <RibbonDivider />
        <Tooltip>
          <TooltipTrigger
            render={
              <Button
                aria-label={isDark ? 'Switch to light mode' : 'Switch to dark mode'}
                className="size-7 sm:size-7"
                onClick={() => setTheme(isDark ? 'light' : 'dark')}
                size="icon-sm"
                variant="ghost"
              />
            }
          >
            {isDark ? <Sun /> : <Moon />}
          </TooltipTrigger>
          <TooltipPopup side="bottom">{isDark ? 'Light mode' : 'Dark mode'}</TooltipPopup>
        </Tooltip>
      </div>
    </header>
  );
}

/* -------------------------------------------------------------------------- */
/*                                   Ribbon                                   */
/* -------------------------------------------------------------------------- */

function SlideNavigation() {
  const { controller, slide, slideCount } = usePlayground();
  return (
    <>
      <RibbonButton
        disabled={slide === 0}
        label="Previous slide"
        onClick={() => void controller?.previous()}
        shortcut="←"
      >
        <ChevronLeft />
      </RibbonButton>
      <span className="min-w-14 text-center text-[13px] tabular-nums text-foreground/80">
        {slideCount ? slide + 1 : 0} / {slideCount}
      </span>
      <RibbonButton
        disabled={slide >= slideCount - 1}
        label="Next slide"
        onClick={() => void controller?.next()}
        shortcut="→"
      >
        <ChevronRight />
      </RibbonButton>
    </>
  );
}

function HomeRibbon() {
  const p = usePlayground();
  const {
    activeResult,
    blocksOpen,
    parsed,
    resetSample,
    results,
    search,
    setBlocksOpen,
    setSearch,
    stepResult,
  } = usePlayground();

  return (
    <>
      <RibbonGroup label="Slides">
        <RibbonButton label="Reset" onClick={resetSample} showLabel>
          <RotateCcw />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="Navigate">
        <SlideNavigation />
      </RibbonGroup>
      <RibbonGroup label="Editing">
        <RibbonButton
          label="Undo"
          disabled={p.readOnly || p.editingBusy || !p.canUndo}
          onClick={p.undo}
        >
          <Undo2 />
        </RibbonButton>
        <RibbonButton
          label="Redo"
          disabled={p.readOnly || p.editingBusy || !p.canRedo}
          onClick={p.redo}
        >
          <Redo2 />
        </RibbonButton>
        <InputGroup className="h-7 w-60 rounded-md text-[13px] before:rounded-[calc(var(--radius-md)-1px)] sm:text-[13px]">
          <InputGroupAddon>
            <Search className="size-3.5!" />
          </InputGroupAddon>
          <InputGroupInput
            aria-label="Find slide text"
            onChange={(event) => setSearch(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') stepResult(event.shiftKey ? -1 : 1);
            }}
            placeholder="Find"
            size="sm"
            value={search}
          />
          {search ? (
            <InputGroupAddon align="inline-end">
              <span className="whitespace-nowrap text-[11px] tabular-nums text-muted-foreground">
                {results.length === 0
                  ? 'No results'
                  : activeResult === null
                    ? `${results.length} ${results.length === 1 ? 'result' : 'results'}`
                    : `${activeResult + 1} of ${results.length}`}
              </span>
            </InputGroupAddon>
          ) : null}
        </InputGroup>
        <RibbonButton
          disabled={!results.length}
          label="Previous result"
          onClick={() => stepResult(-1)}
          shortcut="Shift+Enter"
        >
          <ChevronUp />
        </RibbonButton>
        <RibbonButton
          disabled={!results.length}
          label="Next result"
          onClick={() => stepResult(1)}
          shortcut="Enter"
        >
          <ChevronDown />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="Select">
        <RibbonButton
          active={blocksOpen}
          disabled={!parsed}
          label="Blocks"
          onClick={() => setBlocksOpen(!blocksOpen)}
          showLabel
        >
          <ScanText />
        </RibbonButton>
      </RibbonGroup>
    </>
  );
}

function InsertRibbon() {
  const { fileName, loadSample, openFilePicker } = usePlayground();

  return (
    <>
      <RibbonGroup label="Open">
        <RibbonButton label="Open presentation" onClick={openFilePicker} showLabel>
          <FileUp />
        </RibbonButton>
        <Menu>
          <MenuTrigger render={<RibbonMenuButton icon={<Presentation />} label="Samples" />} />
          <MenuPopup align="start" className="w-60">
            <MenuGroup>
              <MenuGroupLabel>Sample decks</MenuGroupLabel>
              {SAMPLES.map((sample) => (
                <MenuItem key={sample.file} onClick={() => loadSample(sample.file)}>
                  <span className="flex-1">{sample.label}</span>
                  {sample.file === fileName ? (
                    <span className="text-xs text-muted-foreground">Open</span>
                  ) : null}
                </MenuItem>
              ))}
            </MenuGroup>
          </MenuPopup>
        </Menu>
      </RibbonGroup>
      <span className="px-2 text-xs text-muted-foreground">
        Or drop a .ppt or .pptx file anywhere. Files are parsed locally in your browser.{' '}
        <a
          className="underline underline-offset-2"
          href="/samples/ATTRIBUTION.txt"
          target="_blank"
          rel="noreferrer"
        >
          Sample credits
        </a>
      </span>
    </>
  );
}

const THEME_COLOR_KEYS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const;

function describeSlideSize(widthEmu: number, heightEmu: number) {
  const ratio = widthEmu / heightEmu;
  const name =
    Math.abs(ratio - 16 / 9) < 0.01
      ? 'Widescreen (16:9)'
      : Math.abs(ratio - 4 / 3) < 0.01
        ? 'Standard (4:3)'
        : Math.abs(ratio - 16 / 10) < 0.01
          ? 'Widescreen (16:10)'
          : 'Custom';
  const inches = (emu: number) => (emu / EMU_PER_INCH).toFixed(2).replace(/\.?0+$/, '');
  return { name, dimensions: `${inches(widthEmu)} × ${inches(heightEmu)} in` };
}

function DesignRibbon() {
  const { parsed } = usePlayground();
  const presentation = parsed?.document;
  const theme = presentation?.themes[0];
  const swatches = theme
    ? THEME_COLOR_KEYS.flatMap((key) => {
        const value = theme.colors[key];
        return value ? [{ key, color: value.startsWith('#') ? value : `#${value}` }] : [];
      })
    : [];
  const majorFont = theme?.majorFonts?.latin;
  const minorFont = theme?.minorFonts?.latin;

  if (!presentation) {
    return (
      <span className="px-2 text-xs text-muted-foreground">
        Open a presentation to see its theme.
      </span>
    );
  }

  const size = describeSlideSize(presentation.size.widthEmu, presentation.size.heightEmu);

  return (
    <>
      <RibbonGroup className="gap-2 px-2 text-[13px]" label="Themes">
        <span className="text-muted-foreground">Theme</span>
        <span className="font-medium">{theme?.name ?? 'Untitled theme'}</span>
        {swatches.length ? (
          <span aria-label="Theme colors" className="flex items-center gap-0.5" role="list">
            {swatches.map((swatch) => (
              <Tooltip key={swatch.key}>
                <TooltipTrigger
                  render={
                    <span
                      aria-label={`${swatch.key} ${swatch.color}`}
                      className="size-4 rounded-[3px] border border-black/10 dark:border-white/15"
                      role="listitem"
                      style={{ backgroundColor: swatch.color }}
                    />
                  }
                />
                <TooltipPopup side="bottom">
                  {swatch.key} · {swatch.color.toUpperCase()}
                </TooltipPopup>
              </Tooltip>
            ))}
          </span>
        ) : null}
      </RibbonGroup>
      {majorFont || minorFont ? (
        <RibbonGroup className="gap-2 px-2 text-[13px]" label="Fonts">
          <span className="font-serif text-base leading-none">Aa</span>
          <span>
            {majorFont ?? '—'}
            <span className="text-muted-foreground"> / {minorFont ?? '—'}</span>
          </span>
        </RibbonGroup>
      ) : null}
      <RibbonGroup className="gap-2 px-2 text-[13px]" label="Slide size">
        <RectangleHorizontal className="size-4 text-foreground/70" />
        <span>{size.name}</span>
        <span className="text-muted-foreground">{size.dimensions}</span>
      </RibbonGroup>
    </>
  );
}

function SlideShowRibbon() {
  const { parsed, startSlideShow } = usePlayground();

  return (
    <>
      <RibbonGroup label="Start slide show">
        <RibbonButton
          disabled={!parsed}
          label="From Beginning"
          onClick={() => startSlideShow(true)}
          shortcut="F5"
          showLabel
        >
          <Play />
        </RibbonButton>
        <RibbonButton
          disabled={!parsed}
          label="From Current Slide"
          onClick={() => startSlideShow(false)}
          shortcut="Shift+F5"
          showLabel
        >
          <MonitorPlay />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="Navigate">
        <SlideNavigation />
      </RibbonGroup>
      <span className="px-2 text-xs text-muted-foreground">
        Use ← → or Space to advance. Esc ends the show.
      </span>
    </>
  );
}

function ZoomMenu() {
  const { resolvedZoom, setZoom, zoom } = usePlayground();
  return (
    <Menu>
      <MenuTrigger
        render={
          <RibbonMenuButton
            icon={<Search />}
            label={zoom === 'fit-page' ? 'Fit' : `${Math.round(resolvedZoom)}%`}
          />
        }
      />
      <MenuPopup align="start" className="w-44">
        <MenuGroup>
          <MenuGroupLabel>Zoom</MenuGroupLabel>
          {ZOOM_PRESETS.map((value) => (
            <MenuItem key={value} onClick={() => setZoom(value)}>
              {value}%
            </MenuItem>
          ))}
        </MenuGroup>
        <MenuSeparator />
        <MenuItem onClick={() => setZoom('fit-page')}>Fit slide to window</MenuItem>
        <MenuItem onClick={() => setZoom('fit-width')}>Fit to width</MenuItem>
      </MenuPopup>
    </Menu>
  );
}

function ViewRibbon() {
  const p = usePlayground();

  return (
    <>
      <RibbonGroup label="Presentation views">
        <RibbonButton
          active={p.mode === 'slide'}
          label="Normal"
          onClick={() => p.setMode('slide')}
          showLabel
        >
          <RectangleHorizontal />
        </RibbonButton>
        <RibbonButton
          active={p.mode === 'continuous'}
          label="Continuous"
          onClick={() => p.setMode('continuous')}
          showLabel
        >
          <GalleryVertical />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="Show">
        <RibbonToggle
          checked={p.thumbnailsOpen}
          label="Thumbnails"
          onCheckedChange={p.setThumbnailsOpen}
        />
        <RibbonToggle checked={p.notesOpen} label="Notes" onCheckedChange={p.setNotesOpen} />
        <RibbonToggle
          checked={p.slideLabels}
          label="Slide labels"
          onCheckedChange={p.setSlideLabels}
        />
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <ZoomMenu />
        <RibbonButton
          active={p.zoom === 'fit-page'}
          label="Fit to Window"
          onClick={() => p.setZoom('fit-page')}
          showLabel
        >
          <Scan />
        </RibbonButton>
        <RibbonButton
          disabled={p.resolvedZoom <= MIN_ZOOM}
          label="Zoom out"
          onClick={() => p.setZoom(clampZoom(p.resolvedZoom - ZOOM_STEP))}
        >
          <Minus />
        </RibbonButton>
        <RibbonButton
          disabled={p.resolvedZoom >= MAX_ZOOM}
          label="Zoom in"
          onClick={() => p.setZoom(clampZoom(p.resolvedZoom + ZOOM_STEP))}
        >
          <Plus />
        </RibbonButton>
      </RibbonGroup>
    </>
  );
}

function DeveloperRibbon() {
  const p = usePlayground();
  const presentation = p.parsed?.document;

  return (
    <>
      <RibbonGroup label="Inspect">
        <RibbonToggle
          checked={p.diagnostics}
          label="Diagnostics"
          onCheckedChange={p.setDiagnostics}
        />
        <RibbonToggle
          checked={p.blocksOpen}
          disabled={!p.parsed}
          label="Block inspector"
          onCheckedChange={p.setBlocksOpen}
        />
      </RibbonGroup>
      <RibbonGroup className="gap-4 px-3 text-xs" label="Runtime">
        <span className="flex items-center gap-1.5">
          <span
            className={cn(
              'size-1.5 rounded-full',
              presentation ? 'bg-emerald-500' : 'bg-amber-500',
            )}
          />
          <span className="text-muted-foreground">Parser</span>
          <span className="font-medium">{presentation ? 'WASM · ready' : 'Loading'}</span>
        </span>
        <span className="flex items-center gap-1.5">
          <span className="text-muted-foreground">Format</span>
          <span className="font-medium">{presentation?.format.toUpperCase() ?? '—'}</span>
        </span>
        <span className="flex items-center gap-1.5">
          <span className="text-muted-foreground">Parse warnings</span>
          <span className="font-medium tabular-nums">{p.parsed?.warnings.length ?? 0}</span>
        </span>
        <span className="text-muted-foreground">
          Virtualized · {VIRTUALIZATION.initialSlides} initial slides, batches of{' '}
          {VIRTUALIZATION.batchSize}
        </span>
      </RibbonGroup>
    </>
  );
}

function Ribbon() {
  const { readOnly, setReadOnly, legacy } = usePlayground();
  const [tab, setTab] = React.useState<RibbonTab>('Home');
  const panelClass = 'flex items-center';

  return (
    <Tabs
      className="shrink-0 gap-0 px-2 pb-2"
      onValueChange={(value) => setTab(value as RibbonTab)}
      value={tab}
    >
      <div className="flex min-w-0 items-center gap-2">
        <div className="min-w-0 overflow-x-auto [scrollbar-width:none]">
          <TabsList
            className="h-8 gap-0.5 p-0 *:data-[slot=tab-indicator]:bg-brand *:data-[slot=tabs-tab]:hover:bg-black/5 dark:*:data-[slot=tabs-tab]:hover:bg-white/5"
            variant="underline"
          >
            {RIBBON_TABS.map((name) => (
              <TabsTab
                className="h-7 grow-0 rounded-sm px-2.5 text-[13px] font-normal text-foreground/75 hover:text-foreground data-active:font-semibold data-active:text-brand"
                key={name}
                value={name}
              >
                {name}
              </TabsTab>
            ))}
          </TabsList>
        </div>
        <div className="ms-auto flex shrink-0 items-center">
          <RibbonToggle
            checked={readOnly}
            disabled={legacy}
            label="Read only"
            onCheckedChange={setReadOnly}
          />
        </div>
      </div>
      <div className="mt-1 flex h-11 items-center overflow-x-auto rounded-lg border bg-background px-2 shadow-xs/5 [scrollbar-width:none]">
        <TabsPanel className={panelClass} value="Home">
          <HomeRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Insert">
          <InsertRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Design">
          <DesignRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Slide Show">
          <SlideShowRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="View">
          <ViewRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Developer">
          <DeveloperRibbon />
        </TabsPanel>
      </div>
    </Tabs>
  );
}

/* -------------------------------------------------------------------------- */
/*                                 Slide rail                                 */
/* -------------------------------------------------------------------------- */

function SlideRail() {
  const { controller, slide, slideCount } = usePlayground();
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const virtualizer = useVirtualizer({
    count: slideCount,
    enabled: slideCount > 0,
    estimateSize: () => THUMBNAIL_ROW_ESTIMATE,
    getScrollElement: () => scrollRef.current,
    initialRect: { height: 640, width: THUMBNAIL_WIDTH + 40 },
    overscan: THUMBNAIL_OVERSCAN,
    paddingStart: 6,
    paddingEnd: 12,
    useFlushSync: false,
  });
  const virtualThumbnails = virtualizer.getVirtualItems();
  const visibleKey = virtualThumbnails.map((thumbnail) => thumbnail.index).join(',');
  const renderWindow = React.useMemo(
    () =>
      buildThumbnailRenderWindow(
        visibleKey ? visibleKey.split(',').map((index) => Number(index)) : [],
        slideCount,
      ),
    [slideCount, visibleKey],
  );
  const { thumbnails } = usePptxViewerThumbnails(controller, {
    renderWindow,
    resolution: { maxWidth: THUMBNAIL_WIDTH, maxHeight: 126 },
  });

  React.useEffect(() => {
    if (slideCount === 0) return;
    virtualizer.scrollToIndex(slide, { align: 'auto' });
  }, [slide, slideCount, virtualizer]);

  return (
    <nav
      aria-label="Slide thumbnails"
      className="flex w-[208px] shrink-0 flex-col"
      id="playground-slide-thumbnails"
    >
      <div ref={scrollRef} className="min-h-0 flex-1 overflow-y-auto [scrollbar-width:thin]">
        {thumbnails.length ? (
          <div className="relative w-full" style={{ height: virtualizer.getTotalSize() }}>
            {virtualThumbnails.map((virtualThumbnail) => {
              const thumbnail = thumbnails[virtualThumbnail.index];
              if (!thumbnail) return null;
              const active = thumbnail.slideIndex === slide;
              return (
                <div
                  key={virtualThumbnail.key}
                  ref={virtualizer.measureElement}
                  className="absolute inset-x-0 top-0 pb-3"
                  data-index={virtualThumbnail.index}
                  style={{ transform: `translateY(${virtualThumbnail.start}px)` }}
                >
                  <button
                    aria-current={active ? 'page' : undefined}
                    aria-label={`Go to slide ${thumbnail.slideNumber}`}
                    aria-posinset={thumbnail.slideNumber}
                    aria-setsize={slideCount}
                    className="group flex w-full items-start gap-2 ps-1 pe-3 text-left outline-none"
                    data-status={thumbnail.status}
                    onClick={() => void controller?.goToSlide(thumbnail.slideIndex)}
                    type="button"
                  >
                    <span
                      className={cn(
                        'w-5 shrink-0 text-end text-[11px] tabular-nums text-foreground/60',
                        active && 'font-semibold text-brand',
                      )}
                    >
                      {thumbnail.slideNumber}
                    </span>
                    <span
                      className={cn(
                        'slide-thumb-preview block overflow-hidden rounded-[3px] bg-white ring-1 ring-black/10 transition-shadow group-hover:ring-2 group-hover:ring-foreground/25 group-focus-visible:ring-2 group-focus-visible:ring-ring dark:ring-white/10',
                        active && 'ring-2 ring-brand! ring-offset-2 ring-offset-chrome',
                      )}
                      ref={thumbnail.containerRef}
                      style={{ width: thumbnail.width, height: thumbnail.height }}
                    />
                  </button>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="flex flex-col gap-3 ps-8 pe-3 pt-1.5">
            {[0, 1, 2, 3].map((index) => (
              <div
                key={index}
                className="aspect-video w-full animate-pulse rounded-[3px] bg-foreground/6"
              />
            ))}
          </div>
        )}
      </div>
    </nav>
  );
}

/* -------------------------------------------------------------------------- */
/*                                 Status bar                                 */
/* -------------------------------------------------------------------------- */

const statusButton =
  'inline-flex h-5 items-center gap-1 rounded-sm px-1.5 text-foreground/70 outline-none transition-colors hover:bg-black/5 hover:text-foreground focus-visible:ring-2 focus-visible:ring-ring disabled:pointer-events-none disabled:opacity-50 dark:hover:bg-white/8 [&_svg]:size-3.5';

function StatusButton({
  active,
  children,
  className,
  label,
  ...props
}: React.ComponentProps<'button'> & { active?: boolean; label: string }) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <button
            aria-label={label}
            aria-pressed={active}
            className={cn(
              statusButton,
              active && 'bg-black/6 text-foreground dark:bg-white/10',
              className,
            )}
            type="button"
            {...props}
          />
        }
      >
        {children}
      </TooltipTrigger>
      <TooltipPopup side="top">{label}</TooltipPopup>
    </Tooltip>
  );
}

function StatusBar() {
  const p = usePlayground();
  const hasDeck = Boolean(p.parsed);
  const zoomValue = clampZoom(p.resolvedZoom);

  return (
    <footer className="flex h-7 shrink-0 items-center gap-4 bg-chrome px-3 text-[11px] text-foreground/70">
      <span className="tabular-nums">
        {hasDeck ? `Slide ${p.slide + 1} of ${p.slideCount}` : 'Opening presentation…'}
      </span>
      {hasDeck ? (
        <span className="hidden sm:inline">{p.parsed?.document.format.toUpperCase()}</span>
      ) : null}
      <span className="flex items-center gap-1.5">
        {hasDeck ? null : <Spinner className="size-3" />}
        {hasDeck ? 'Ready' : 'Loading…'}
      </span>
      <div className="flex-1" />
      <div className="flex items-center gap-0.5">
        <StatusButton
          active={p.notesOpen}
          label="Speaker notes"
          onClick={() => p.setNotesOpen(!p.notesOpen)}
        >
          <StickyNote />
          Notes
        </StatusButton>
        <StatusButton
          active={p.thumbnailsOpen}
          label="Thumbnails"
          onClick={() => p.setThumbnailsOpen(!p.thumbnailsOpen)}
        >
          <PanelLeft />
        </StatusButton>
      </div>
      <div className="flex items-center gap-0.5">
        <StatusButton active={p.mode === 'slide'} label="Normal" onClick={() => p.setMode('slide')}>
          <RectangleHorizontal />
        </StatusButton>
        <StatusButton
          active={p.mode === 'continuous'}
          label="Continuous"
          onClick={() => p.setMode('continuous')}
        >
          <GalleryVertical />
        </StatusButton>
        <StatusButton
          disabled={!hasDeck}
          label="Slide Show"
          onClick={() => p.startSlideShow(false)}
        >
          <MonitorPlay />
        </StatusButton>
      </div>
      <div className="flex items-center gap-1">
        <StatusButton
          className="w-5 justify-center px-0"
          disabled={zoomValue <= MIN_ZOOM}
          label="Zoom out"
          onClick={() => p.setZoom(clampZoom(zoomValue - ZOOM_STEP))}
        >
          <Minus className="size-3!" />
        </StatusButton>
        <Slider
          aria-label="Zoom"
          className="w-24 **:data-[slot=slider-control]:min-w-0 **:data-[slot=slider-indicator]:bg-foreground/40 **:data-[slot=slider-thumb]:size-3 sm:**:data-[slot=slider-thumb]:size-3"
          max={MAX_ZOOM}
          min={MIN_ZOOM}
          onValueChange={(value) =>
            p.setZoom(clampZoom(Array.isArray(value) ? (value[0] ?? 100) : value))
          }
          value={zoomValue}
        />
        <StatusButton
          className="w-5 justify-center px-0"
          disabled={zoomValue >= MAX_ZOOM}
          label="Zoom in"
          onClick={() => p.setZoom(clampZoom(zoomValue + ZOOM_STEP))}
        >
          <Plus className="size-3!" />
        </StatusButton>
        <StatusButton
          className="w-10 justify-end tabular-nums"
          label="Reset zoom to 100%"
          onClick={() => p.setZoom(100)}
        >
          {zoomValue}%
        </StatusButton>
        <StatusButton
          active={p.zoom === 'fit-page'}
          label="Fit slide to current window"
          onClick={() => p.setZoom('fit-page')}
        >
          <Scan />
        </StatusButton>
      </div>
    </footer>
  );
}

/* -------------------------------------------------------------------------- */
/*                                     App                                    */
/* -------------------------------------------------------------------------- */

type OpenedFile = { name: string; source: string | File };

export function App() {
  const viewer = usePptxViewer();
  const inputRef = React.useRef<HTMLInputElement>(null);
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const blockRootRef = React.useRef<HTMLDivElement>(null);
  const dragDepthRef = React.useRef(0);
  const restoreAfterShowRef = React.useRef<{ mode: ViewerMode; zoom: ViewerZoomLevel } | null>(
    null,
  );
  const [opened, setOpened] = React.useState<OpenedFile>({
    name: DEFAULT_SAMPLE.file,
    source: DEFAULT_SAMPLE.source,
  });
  const [source, setSource] = React.useState<PresentationSource>(DEFAULT_SAMPLE.source);
  const [mode, setMode] = React.useState<ViewerMode>('slide');
  const [zoom, setZoom] = React.useState<ViewerZoomLevel>('fit-page');
  const [resolvedZoom, setResolvedZoom] = React.useState(100);
  const [slide, setSlide] = React.useState(0);
  const [slideCount, setSlideCount] = React.useState(0);
  const [search, setSearchValue] = React.useState('');
  const [results, setResults] = React.useState<readonly PresentationSearchResult[]>([]);
  const [activeResult, setActiveResult] = React.useState<number | null>(null);
  const [diagnostics, setDiagnostics] = React.useState(false);
  const [thumbnailsOpen, setThumbnailsOpen] = React.useState(true);
  const [blocksOpen, setBlocksOpen] = React.useState(false);
  const [notesOpen, setNotesOpen] = React.useState(false);
  const [slideLabels, setSlideLabels] = React.useState(false);
  const [presenting, setPresenting] = React.useState(false);
  const [dragging, setDragging] = React.useState(false);
  const [parsed, setParsed] = React.useState<ParsedPresentation | null>(null);
  const [readOnly, setReadOnly] = React.useState(false);
  const pendingEditRenderRef = React.useRef<(() => void) | null>(null);
  const finishEditRender = React.useCallback(() => {
    pendingEditRenderRef.current?.();
    pendingEditRenderRef.current = null;
  }, []);
  const applyEditedPresentation = React.useCallback(
    (next: ParsedPresentation) =>
      new Promise<void>((resolve) => {
        pendingEditRenderRef.current?.();
        pendingEditRenderRef.current = resolve;
        setSource(next);
      }),
    [],
  );
  const editing = usePptxEditing(opened, applyEditedPresentation);
  React.useEffect(() => finishEditRender, [finishEditRender]);
  const legacy = parsed?.document.format === 'ppt';

  const open = React.useCallback(
    (next: OpenedFile, cacheBust = false) => {
      finishEditRender();
      setParsed(null);
      setSlideCount(0);
      setSlide(0);
      setSearchValue('');
      setActiveResult(null);
      setOpened(next);
      setSource(
        typeof next.source === 'string' && cacheBust
          ? `${next.source}?t=${Date.now()}`
          : next.source,
      );
    },
    [finishEditRender],
  );

  const openFile = React.useCallback(
    (file?: File) => {
      if (!file) return;
      const extension = file.name.split('.').pop()?.toLowerCase();
      if (extension !== 'ppt' && extension !== 'pptx') return;
      open({ name: file.name, source: file });
    },
    [open],
  );

  const stepResult = React.useCallback(
    (direction: 1 | -1) => {
      if (!results.length) return;
      setActiveResult(
        (current) =>
          ((current ?? (direction > 0 ? -1 : 0)) + direction + results.length) % results.length,
      );
    },
    [results.length],
  );

  const startSlideShow = React.useCallback(
    (fromBeginning: boolean) => {
      const element = canvasRef.current;
      if (!element || document.fullscreenElement) return;
      restoreAfterShowRef.current = { mode, zoom };
      setMode('slide');
      setZoom('fit-page');
      if (fromBeginning) void viewer.controller?.goToSlide(0);
      element.requestFullscreen().catch(() => {
        const restore = restoreAfterShowRef.current;
        restoreAfterShowRef.current = null;
        if (restore) {
          setMode(restore.mode);
          setZoom(restore.zoom);
        }
      });
    },
    [mode, viewer.controller, zoom],
  );

  React.useEffect(() => {
    const onFullscreenChange = () => {
      const active = document.fullscreenElement === canvasRef.current;
      setPresenting(active);
      if (!active && restoreAfterShowRef.current) {
        const restore = restoreAfterShowRef.current;
        restoreAfterShowRef.current = null;
        setMode(restore.mode);
        setZoom(restore.zoom);
      }
    };
    document.addEventListener('fullscreenchange', onFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', onFullscreenChange);
  }, []);

  React.useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        event.target instanceof HTMLInputElement ||
        event.target instanceof HTMLTextAreaElement ||
        (event.target instanceof HTMLElement && event.target.isContentEditable)
      )
        return;
      if (
        (event.metaKey || event.ctrlKey) &&
        ['z', 'y'].includes(event.key.toLowerCase()) &&
        !readOnly &&
        !legacy &&
        !presenting
      ) {
        event.preventDefault();
        void editing.update(event.key.toLowerCase() === 'y' || event.shiftKey ? 'redo' : 'undo');
        return;
      }
      if (event.key === 'F5') {
        event.preventDefault();
        startSlideShow(!event.shiftKey);
        return;
      }
      const forward = event.key === 'ArrowRight' || event.key === 'PageDown';
      const back = event.key === 'ArrowLeft' || event.key === 'PageUp';
      if (forward || (presenting && event.key === ' ')) {
        event.preventDefault();
        void viewer.controller?.next();
      }
      if (back) void viewer.controller?.previous();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [presenting, startSlideShow, viewer.controller, readOnly, legacy, editing.update]);

  const playground: Playground = {
    controller: viewer.controller,
    parsed,
    fileName: opened.name,
    slide,
    slideCount,
    mode,
    setMode,
    zoom,
    setZoom,
    resolvedZoom,
    search,
    setSearch: (value) => {
      setSearchValue(value);
      setActiveResult(null);
    },
    results,
    activeResult,
    stepResult,
    diagnostics,
    setDiagnostics,
    thumbnailsOpen,
    setThumbnailsOpen,
    blocksOpen,
    setBlocksOpen,
    notesOpen,
    setNotesOpen,
    slideLabels,
    setSlideLabels,
    openFilePicker: () => inputRef.current?.click(),
    loadSample: (file) => {
      const sample = SAMPLES.find((candidate) => candidate.file === file);
      if (sample) open({ name: sample.file, source: sample.source }, true);
    },
    resetSample: () => open({ name: DEFAULT_SAMPLE.file, source: DEFAULT_SAMPLE.source }, true),
    downloadOriginal: () => {
      editing.download(opened.name);
    },
    readOnly: readOnly || legacy,
    setReadOnly,
    legacy,
    editingBusy: editing.busy,
    editingReady: editing.ready,
    canUndo: editing.canUndo,
    canRedo: editing.canRedo,
    undo: () => {
      void editing.update('undo');
    },
    redo: () => {
      void editing.update('redo');
    },
    startSlideShow,
  };

  const isFileDrag = (event: React.DragEvent) => event.dataTransfer.types.includes('Files');

  return (
    <PlaygroundContext.Provider value={playground}>
      <div
        className="relative flex h-dvh flex-col overflow-hidden bg-chrome text-foreground"
        onDragEnter={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          dragDepthRef.current += 1;
          setDragging(true);
        }}
        onDragLeave={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
          if (dragDepthRef.current === 0) setDragging(false);
        }}
        onDragOver={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          event.dataTransfer.dropEffect = 'copy';
        }}
        onDrop={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          dragDepthRef.current = 0;
          setDragging(false);
          openFile(event.dataTransfer.files[0]);
        }}
      >
        <input
          ref={inputRef}
          accept=".ppt,.pptx,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          className="hidden"
          onChange={(event) => {
            const file = event.target.files?.[0];
            event.target.value = '';
            openFile(file);
          }}
          type="file"
        />

        <TitleBar />
        <Ribbon />

        {editing.error && (
          <div role="alert" className="px-3 pb-2 text-sm text-destructive-foreground">
            {editing.error}
          </div>
        )}
        {!readOnly && (
          <div
            className="h-6 shrink-0 truncate px-3 pb-2 text-xs text-muted-foreground"
            role="status"
          >
            {legacy
              ? 'Legacy PPT files are read only. Open a PPTX file to edit.'
              : editing.busy
                ? 'Saving changes…'
                : 'Click text to edit. Drag an object’s border to move it. Use the handles to resize. Download saves your changes.'}
          </div>
        )}

        <main className="flex min-h-0 flex-1 gap-2 px-2" id="viewer-workspace">
          {thumbnailsOpen ? <SlideRail /> : null}
          <div
            ref={canvasRef}
            className={cn(
              'slide-canvas relative min-w-0 flex-1 overflow-hidden rounded-lg border bg-canvas',
              presenting && 'rounded-none border-0 [--canvas:#000]',
            )}
            data-presenting={presenting || undefined}
            data-mode={mode}
          >
            <div ref={blockRootRef} className="h-full min-h-0">
              <ReactPptxViewer
                ref={viewer.ref}
                source={source}
                mode={mode}
                zoom={zoom}
                slideIndex={slide}
                showDiagnostics={diagnostics && !presenting}
                showNotes={notesOpen && !presenting}
                showSlideLabels={slideLabels && !presenting}
                virtualization={VIRTUALIZATION}
                searchQuery={search}
                activeSearchResult={activeResult}
                onLoad={(next) => {
                  setParsed(next);
                  setSlideCount(next.document.slides.length);
                }}
                onSlideChange={setSlide}
                onReady={finishEditRender}
                onSearchResults={setResults}
                onZoomChange={(state) => setResolvedZoom(state.resolvedZoom)}
                onError={(error) => {
                  finishEditRender();
                  console.error(error);
                }}
                height="100%"
              />
            </div>
          </div>
          <ObjectEditor
            key={
              opened.name +
              String(typeof opened.source === 'string' ? opened.source : opened.source.lastModified)
            }
            rootRef={blockRootRef}
            parsed={parsed}
            disabled={readOnly || legacy || presenting || !editing.ready}
            busy={editing.busy}
            revision={editing.revision}
            slideIndex={mode === 'slide' ? slide : undefined}
            onEdit={editing.update}
          />
          {blocksOpen && parsed ? (
            <BlockInspector
              key={parsed.document.documentId ?? opened.name}
              model={parsed.document}
              rootRef={blockRootRef}
              controller={viewer.controller}
              onClose={() => setBlocksOpen(false)}
            />
          ) : null}
        </main>

        <StatusBar />

        {dragging ? (
          <div className="pointer-events-none absolute inset-0 z-30 flex items-center justify-center bg-background/70 backdrop-blur-[2px]">
            <div className="flex flex-col items-center gap-3 rounded-xl border-2 border-dashed border-brand/60 bg-background px-10 py-8 text-center shadow-lg">
              <Upload className="size-6 text-brand" />
              <div className="text-sm font-medium">Drop to open presentation</div>
              <div className="text-xs text-muted-foreground">
                PPT and PPTX files are parsed locally in your browser.
              </div>
            </div>
          </div>
        ) : null}
      </div>
    </PlaygroundContext.Provider>
  );
}
