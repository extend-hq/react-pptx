/**
 * Bounded browser renderer for the GDI records commonly embedded in PowerPoint
 * EMF and WMF preview images. The parser is intentionally local to the viewer:
 * metafile bytes never leave the browser and malformed records fail closed.
 */

export interface MetafileRenderOptions {
  maxWidth?: number;
  maxHeight?: number;
  dpiScale?: number;
  maxCanvasDimension?: number;
  maxRecords?: number;
  fontFamilyMap?: Readonly<Record<string, string>>;
}

type RenderCanvas = HTMLCanvasElement | OffscreenCanvas;

interface Point {
  x: number;
  y: number;
}

interface Bounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

interface Matrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

interface Pen {
  kind: 'pen';
  color: string;
  width: number;
  style: number;
  null: boolean;
}

interface Brush {
  kind: 'brush';
  color: string;
  style: number;
  null: boolean;
}

interface Font {
  kind: 'font';
  family: string;
  height: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  escapement: number;
}

type GdiObject = Pen | Brush | Font;

interface DrawingState {
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  backgroundColor: string;
  backgroundMode: number;
  textAlign: number;
  fillMode: CanvasFillRule;
  current: Point;
  windowOrigin: Point;
  windowExtent: Point;
  viewportOrigin: Point;
  viewportExtent: Point;
  mappingEnabled: boolean;
  world: Matrix;
}

interface Surface {
  canvas: RenderCanvas;
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scaleX: number;
  scaleY: number;
}

const DEFAULT_MAX_DIMENSION = 4_096;
const DEFAULT_MAX_RECORDS = 200_000;
const DEFAULT_FONT: Font = {
  kind: 'font',
  family: 'sans-serif',
  height: 12,
  weight: 400,
  italic: false,
  underline: false,
  strike: false,
  escapement: 0,
};
const DEFAULT_PEN: Pen = { kind: 'pen', color: '#000000', width: 1, style: 0, null: false };
const DEFAULT_BRUSH: Brush = {
  kind: 'brush',
  color: '#ffffff',
  style: 0,
  null: false,
};

function finitePositive(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : fallback;
}

function readOptions(
  maxWidth: number | undefined,
  maxHeight: number | undefined,
  optionsOrScale: MetafileRenderOptions | number | undefined,
): Required<Omit<MetafileRenderOptions, 'fontFamilyMap'>> & {
  fontFamilyMap: Readonly<Record<string, string>>;
} {
  const options =
    typeof optionsOrScale === 'number' ? { dpiScale: optionsOrScale } : optionsOrScale;
  return {
    maxWidth: finitePositive(options?.maxWidth ?? maxWidth, Number.POSITIVE_INFINITY),
    maxHeight: finitePositive(options?.maxHeight ?? maxHeight, Number.POSITIVE_INFINITY),
    dpiScale: Math.min(4, finitePositive(options?.dpiScale, 1)),
    maxCanvasDimension: Math.floor(
      finitePositive(options?.maxCanvasDimension, DEFAULT_MAX_DIMENSION),
    ),
    maxRecords: Math.floor(finitePositive(options?.maxRecords, DEFAULT_MAX_RECORDS)),
    fontFamilyMap: options?.fontFamilyMap ?? {},
  };
}

function createSurface(
  logicalWidth: number,
  logicalHeight: number,
  options: ReturnType<typeof readOptions>,
): Surface | undefined {
  if (!(logicalWidth > 0) || !(logicalHeight > 0)) return undefined;
  const fit = Math.min(
    1,
    options.maxWidth / logicalWidth,
    options.maxHeight / logicalHeight,
    options.maxCanvasDimension / (logicalWidth * options.dpiScale),
    options.maxCanvasDimension / (logicalHeight * options.dpiScale),
  );
  const width = Math.max(1, Math.round(logicalWidth * fit * options.dpiScale));
  const height = Math.max(1, Math.round(logicalHeight * fit * options.dpiScale));
  let canvas: RenderCanvas;
  if (typeof OffscreenCanvas !== 'undefined') canvas = new OffscreenCanvas(width, height);
  else if (typeof document !== 'undefined') {
    canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
  } else return undefined;
  const context = canvas.getContext('2d');
  if (!context) return undefined;
  return {
    canvas,
    context: context as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
    scaleX: width / logicalWidth,
    scaleY: height / logicalHeight,
  };
}

async function exportPng(canvas: RenderCanvas): Promise<string | null> {
  if ('convertToBlob' in canvas) {
    const blob = await canvas.convertToBlob({ type: 'image/png' });
    if (typeof FileReader !== 'undefined') {
      return await new Promise<string | null>((resolve) => {
        const reader = new FileReader();
        reader.onerror = () => resolve(null);
        reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : null);
        reader.readAsDataURL(blob);
      });
    }
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return typeof btoa === 'function' ? `data:image/png;base64,${btoa(binary)}` : null;
  }
  return canvas.toDataURL('image/png');
}

function colorRef(view: DataView, offset: number): string {
  const hex = (value: number) => value.toString(16).padStart(2, '0');
  return `#${hex(view.getUint8(offset))}${hex(view.getUint8(offset + 1))}${hex(view.getUint8(offset + 2))}`;
}

function readUtf16(view: DataView, offset: number, characters: number): string {
  const end = Math.min(view.byteLength, offset + characters * 2);
  let value = '';
  for (let cursor = offset; cursor + 1 < end; cursor += 2) {
    const code = view.getUint16(cursor, true);
    if (code === 0) break;
    value += String.fromCharCode(code);
  }
  return value;
}

function readAnsi(view: DataView, offset: number, length: number): string {
  const end = Math.min(view.byteLength, offset + length);
  let value = '';
  for (let cursor = offset; cursor < end; cursor += 1) {
    const code = view.getUint8(cursor);
    if (code === 0) break;
    value += String.fromCharCode(code);
  }
  return value;
}

function fontFamily(face: string, map: Readonly<Record<string, string>>): string {
  const clean = face.trim() || 'sans-serif';
  const mapped = map[clean] ?? map[clean.toLowerCase()] ?? clean;
  return /^(?:serif|sans-serif|monospace|system-ui)$/i.test(mapped)
    ? mapped
    : `"${mapped.replaceAll('"', '\\"')}"`;
}

function cloneState(state: DrawingState): DrawingState {
  return {
    ...state,
    pen: { ...state.pen },
    brush: { ...state.brush },
    font: { ...state.font },
    current: { ...state.current },
    windowOrigin: { ...state.windowOrigin },
    windowExtent: { ...state.windowExtent },
    viewportOrigin: { ...state.viewportOrigin },
    viewportExtent: { ...state.viewportExtent },
    world: { ...state.world },
  };
}

function initialState(bounds: Bounds): DrawingState {
  return {
    pen: { ...DEFAULT_PEN },
    brush: { ...DEFAULT_BRUSH },
    font: { ...DEFAULT_FONT },
    textColor: '#000000',
    backgroundColor: '#ffffff',
    backgroundMode: 2,
    textAlign: 0,
    fillMode: 'nonzero',
    current: { x: 0, y: 0 },
    windowOrigin: { x: bounds.left, y: bounds.top },
    windowExtent: { x: bounds.right - bounds.left, y: bounds.bottom - bounds.top },
    viewportOrigin: { x: 0, y: 0 },
    viewportExtent: { x: bounds.right - bounds.left, y: bounds.bottom - bounds.top },
    mappingEnabled: false,
    world: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
  };
}

class GdiRenderer {
  readonly objects = new Map<number, GdiObject>();
  readonly stack: DrawingState[] = [];
  state: DrawingState;

  constructor(
    readonly context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
    readonly bounds: Bounds,
    readonly scaleX: number,
    readonly scaleY: number,
    readonly familyMap: Readonly<Record<string, string>>,
  ) {
    this.state = initialState(bounds);
  }

  point(x: number, y: number): Point {
    const matrix = this.state.world;
    const transformed = {
      x: matrix.a * x + matrix.c * y + matrix.e,
      y: matrix.b * x + matrix.d * y + matrix.f,
    };
    const deviceX = this.state.mappingEnabled
      ? ((transformed.x - this.state.windowOrigin.x) / (this.state.windowExtent.x || 1)) *
          (this.state.viewportExtent.x || 1) +
        this.state.viewportOrigin.x
      : transformed.x;
    const deviceY = this.state.mappingEnabled
      ? ((transformed.y - this.state.windowOrigin.y) / (this.state.windowExtent.y || 1)) *
          (this.state.viewportExtent.y || 1) +
        this.state.viewportOrigin.y
      : transformed.y;
    return {
      x: (deviceX - this.bounds.left) * this.scaleX,
      y: (deviceY - this.bounds.top) * this.scaleY,
    };
  }

  lengthX(value: number): number {
    return (
      Math.abs(
        this.state.mappingEnabled
          ? (value / (this.state.windowExtent.x || 1)) * (this.state.viewportExtent.x || 1)
          : value,
      ) * this.scaleX
    );
  }

  lengthY(value: number): number {
    return (
      Math.abs(
        this.state.mappingEnabled
          ? (value / (this.state.windowExtent.y || 1)) * (this.state.viewportExtent.y || 1)
          : value,
      ) * this.scaleY
    );
  }

  applyPen(): boolean {
    const pen = this.state.pen;
    if (pen.null) return false;
    this.context.strokeStyle = pen.color;
    this.context.lineWidth = Math.max(1, this.lengthX(pen.width || 1));
    const dash = pen.style & 0x0f;
    this.context.setLineDash?.(
      dash === 1 ? [8, 4] : dash === 2 ? [2, 3] : dash === 3 ? [8, 3, 2, 3] : [],
    );
    return true;
  }

  applyBrush(): boolean {
    if (this.state.brush.null) return false;
    this.context.fillStyle = this.state.brush.color;
    return true;
  }

  paintPath(close = false): void {
    if (close) this.context.closePath();
    if (this.applyBrush() && close) this.context.fill(this.state.fillMode);
    if (this.applyPen()) this.context.stroke();
  }

  selectObject(index: number): void {
    const object = index >= 0x80000000 ? stockObject(index & 0x7fffffff) : this.objects.get(index);
    if (!object) return;
    if (object.kind === 'pen') this.state.pen = { ...object };
    else if (object.kind === 'brush') this.state.brush = { ...object };
    else this.state.font = { ...object };
  }

  drawText(text: string, x: number, y: number): void {
    if (!text) return;
    const position = this.point(x, y);
    const size = Math.max(8, this.lengthY(this.state.font.height));
    const weight = Math.max(100, Math.min(900, Math.round(this.state.font.weight / 100) * 100));
    const weightPart = weight === 400 ? '' : `${weight} `;
    this.context.font = `${this.state.font.italic ? 'italic ' : ''}${weightPart}${size}px ${fontFamily(this.state.font.family, this.familyMap)}`;
    this.context.fillStyle = this.state.textColor;
    this.context.textAlign =
      (this.state.textAlign & 6) === 6
        ? 'center'
        : (this.state.textAlign & 2) === 2
          ? 'right'
          : 'left';
    this.context.textBaseline =
      (this.state.textAlign & 24) === 24
        ? 'alphabetic'
        : (this.state.textAlign & 8) === 8
          ? 'bottom'
          : 'top';
    const angle = (-this.state.font.escapement / 10) * (Math.PI / 180);
    if (angle) {
      this.context.save();
      this.context.translate(position.x, position.y);
      this.context.rotate(angle);
      this.context.fillText(text, 0, 0);
      this.context.restore();
    } else this.context.fillText(text, position.x, position.y);
  }
}

function stockObject(index: number): GdiObject | undefined {
  if (index === 0) return { ...DEFAULT_BRUSH };
  if (index === 4) return { ...DEFAULT_BRUSH, color: '#000000' };
  if (index === 5) return { ...DEFAULT_BRUSH, null: true };
  if (index === 6) return { ...DEFAULT_PEN, color: '#ffffff' };
  if (index === 7) return { ...DEFAULT_PEN };
  if (index === 8) return { ...DEFAULT_PEN, null: true };
  if (index >= 10 && index <= 17) return { ...DEFAULT_FONT };
  return undefined;
}

function validRange(view: DataView, offset: number, length: number): boolean {
  return offset >= 0 && length >= 0 && offset <= view.byteLength - length;
}

function drawRectangle(renderer: GdiRenderer, bounds: Bounds, radius?: Point): void {
  const first = renderer.point(bounds.left, bounds.top);
  const second = renderer.point(bounds.right, bounds.bottom);
  const x = Math.min(first.x, second.x);
  const y = Math.min(first.y, second.y);
  const width = Math.abs(second.x - first.x);
  const height = Math.abs(second.y - first.y);
  renderer.context.beginPath();
  if (radius && 'roundRect' in renderer.context) {
    renderer.context.roundRect(x, y, width, height, [
      renderer.lengthX(radius.x) / 2,
      renderer.lengthY(radius.y) / 2,
    ]);
  } else renderer.context.rect(x, y, width, height);
  renderer.paintPath(true);
}

function drawEllipse(renderer: GdiRenderer, bounds: Bounds): void {
  const first = renderer.point(bounds.left, bounds.top);
  const second = renderer.point(bounds.right, bounds.bottom);
  const centerX = (first.x + second.x) / 2;
  const centerY = (first.y + second.y) / 2;
  renderer.context.beginPath();
  renderer.context.ellipse(
    centerX,
    centerY,
    Math.abs(second.x - first.x) / 2,
    Math.abs(second.y - first.y) / 2,
    0,
    0,
    Math.PI * 2,
  );
  renderer.paintPath(true);
}

function drawPoints(renderer: GdiRenderer, points: Point[], close: boolean): void {
  const first = points[0];
  if (!first) return;
  const start = renderer.point(first.x, first.y);
  renderer.context.beginPath();
  renderer.context.moveTo(start.x, start.y);
  for (const point of points.slice(1)) {
    const next = renderer.point(point.x, point.y);
    renderer.context.lineTo(next.x, next.y);
  }
  renderer.paintPath(close);
}

function readBounds32(view: DataView, offset: number): Bounds {
  return {
    left: view.getInt32(offset, true),
    top: view.getInt32(offset + 4, true),
    right: view.getInt32(offset + 8, true),
    bottom: view.getInt32(offset + 12, true),
  };
}

function replayEmfRecord(
  renderer: GdiRenderer,
  view: DataView,
  recordOffset: number,
  type: number,
  size: number,
): boolean {
  const data = recordOffset + 8;
  switch (type) {
    case 9:
      renderer.state.windowExtent = {
        x: view.getInt32(data, true),
        y: view.getInt32(data + 4, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 10:
      renderer.state.windowOrigin = {
        x: view.getInt32(data, true),
        y: view.getInt32(data + 4, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 11:
      renderer.state.viewportExtent = {
        x: view.getInt32(data, true),
        y: view.getInt32(data + 4, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 12:
      renderer.state.viewportOrigin = {
        x: view.getInt32(data, true),
        y: view.getInt32(data + 4, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 14:
      return false;
    case 15: {
      const point = renderer.point(view.getInt32(data, true), view.getInt32(data + 4, true));
      renderer.context.fillStyle = colorRef(view, data + 8);
      renderer.context.fillRect(point.x, point.y, 1, 1);
      break;
    }
    case 17:
      renderer.state.mappingEnabled = view.getInt32(data, true) !== 1;
      break;
    case 18:
      renderer.state.backgroundMode = view.getInt32(data, true);
      break;
    case 19:
      renderer.state.fillMode = view.getInt32(data, true) === 1 ? 'evenodd' : 'nonzero';
      break;
    case 22:
      renderer.state.textAlign = view.getInt32(data, true);
      break;
    case 24:
      renderer.state.textColor = colorRef(view, data);
      break;
    case 25:
      renderer.state.backgroundColor = colorRef(view, data);
      break;
    case 27:
      renderer.state.current = {
        x: view.getInt32(data, true),
        y: view.getInt32(data + 4, true),
      };
      break;
    case 30: {
      const bounds = readBounds32(view, data);
      const first = renderer.point(bounds.left, bounds.top);
      const second = renderer.point(bounds.right, bounds.bottom);
      renderer.context.beginPath();
      renderer.context.rect(first.x, first.y, second.x - first.x, second.y - first.y);
      renderer.context.clip();
      break;
    }
    case 33:
      renderer.stack.push(cloneState(renderer.state));
      renderer.context.save();
      break;
    case 34: {
      const restored = renderer.stack.pop();
      if (restored) {
        renderer.state = restored;
        renderer.context.restore();
      }
      break;
    }
    case 35:
      renderer.state.world = {
        a: view.getFloat32(data, true),
        b: view.getFloat32(data + 4, true),
        c: view.getFloat32(data + 8, true),
        d: view.getFloat32(data + 12, true),
        e: view.getFloat32(data + 16, true),
        f: view.getFloat32(data + 20, true),
      };
      break;
    case 37:
      renderer.selectObject(view.getUint32(data, true));
      break;
    case 38:
      renderer.objects.set(view.getUint32(data, true), {
        kind: 'pen',
        style: view.getUint32(data + 4, true),
        width: Math.abs(view.getInt32(data + 8, true)) || 1,
        color: colorRef(view, data + 16),
        null: false,
      });
      break;
    case 39: {
      const style = view.getUint32(data + 4, true);
      renderer.objects.set(view.getUint32(data, true), {
        kind: 'brush',
        style,
        color: colorRef(view, data + 8),
        null: style === 1,
      });
      break;
    }
    case 40:
      renderer.objects.delete(view.getUint32(data, true));
      break;
    case 42:
      drawEllipse(renderer, readBounds32(view, data));
      break;
    case 43:
      drawRectangle(renderer, readBounds32(view, data));
      break;
    case 44:
      drawRectangle(renderer, readBounds32(view, data), {
        x: view.getInt32(data + 16, true),
        y: view.getInt32(data + 20, true),
      });
      break;
    case 54: {
      const end = { x: view.getInt32(data, true), y: view.getInt32(data + 4, true) };
      drawPoints(renderer, [renderer.state.current, end], false);
      renderer.state.current = end;
      break;
    }
    case 82:
      renderer.objects.set(view.getUint32(data, true), {
        kind: 'font',
        height: Math.abs(view.getInt32(data + 4, true)) || 12,
        escapement: view.getInt32(data + 12, true),
        weight: view.getInt32(data + 20, true) || 400,
        italic: view.getUint8(data + 24) !== 0,
        underline: view.getUint8(data + 25) !== 0,
        strike: view.getUint8(data + 26) !== 0,
        family: readUtf16(view, data + 32, 32) || 'sans-serif',
      });
      break;
    case 84: {
      const count = view.getUint32(data + 36, true);
      const stringOffset = view.getUint32(data + 40, true);
      if (count <= 1_000_000 && stringOffset >= 8 && stringOffset + count * 2 <= size) {
        renderer.drawText(
          readUtf16(view, recordOffset + stringOffset, count),
          view.getInt32(data + 28, true),
          view.getInt32(data + 32, true),
        );
      }
      break;
    }
    case 3:
    case 4:
    case 85:
    case 86:
    case 87: {
      const count = view.getUint32(data + 16, true);
      const short = type >= 85;
      const stride = short ? 4 : 8;
      if (count > 1_000_000 || 20 + count * stride > size - 8) break;
      const points: Point[] = [];
      for (let index = 0; index < count; index += 1) {
        const offset = data + 20 + index * stride;
        points.push({
          x: short ? view.getInt16(offset, true) : view.getInt32(offset, true),
          y: short ? view.getInt16(offset + 2, true) : view.getInt32(offset + 4, true),
        });
      }
      drawPoints(renderer, points, type === 3 || type === 86);
      break;
    }
  }
  return true;
}

function emfBounds(view: DataView): Bounds | undefined {
  if (view.byteLength < 88 || view.getUint32(0, true) !== 1) return undefined;
  const bounds = readBounds32(view, 8);
  if (bounds.right > bounds.left && bounds.bottom > bounds.top) return bounds;
  return undefined;
}

export async function renderEmfToDataUrl(
  buffer: ArrayBuffer,
  maxWidth?: number,
  maxHeight?: number,
  optionsOrScale?: MetafileRenderOptions | number,
): Promise<string | null> {
  try {
    const view = new DataView(buffer);
    const bounds = emfBounds(view);
    if (!bounds) return null;
    const options = readOptions(maxWidth, maxHeight, optionsOrScale);
    const surface = createSurface(bounds.right - bounds.left, bounds.bottom - bounds.top, options);
    if (!surface) return null;
    const renderer = new GdiRenderer(
      surface.context,
      bounds,
      surface.scaleX,
      surface.scaleY,
      options.fontFamilyMap,
    );
    let offset = 0;
    let records = 0;
    while (offset + 8 <= view.byteLength && records < options.maxRecords) {
      const type = view.getUint32(offset, true);
      const size = view.getUint32(offset + 4, true);
      if (size < 8 || size % 4 !== 0 || !validRange(view, offset, size)) return null;
      records += 1;
      const keepGoing = replayEmfRecord(renderer, view, offset, type, size);
      offset += size;
      if (!keepGoing) break;
    }
    return await exportPng(surface.canvas);
  } catch {
    return null;
  }
}

const PLACEABLE_WMF_KEY = 0x9ac6cdd7;

interface WmfHeader {
  recordsOffset: number;
  bounds: Bounds;
}

function wmfHeader(view: DataView): WmfHeader | undefined {
  if (view.byteLength < 18) return undefined;
  if (view.getUint32(0, true) === PLACEABLE_WMF_KEY) {
    if (view.byteLength < 40) return undefined;
    const bounds = {
      left: view.getInt16(6, true),
      top: view.getInt16(8, true),
      right: view.getInt16(10, true),
      bottom: view.getInt16(12, true),
    };
    if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) return undefined;
    return { recordsOffset: 40, bounds };
  }
  // A non-placeable WMF has no physical bounds. Start with its logical default;
  // SETWINDOWORG/SETWINDOWEXT records refine the mapping during replay.
  return { recordsOffset: 18, bounds: { left: 0, top: 0, right: 1_024, bottom: 768 } };
}

function addWmfObject(renderer: GdiRenderer, object: GdiObject): void {
  let index = 0;
  while (renderer.objects.has(index)) index += 1;
  renderer.objects.set(index, object);
}

function wmfPoint(view: DataView, offset: number): Point {
  return { x: view.getInt16(offset, true), y: view.getInt16(offset + 2, true) };
}

function replayWmfRecord(
  renderer: GdiRenderer,
  view: DataView,
  data: number,
  functionId: number,
  dataLength: number,
): boolean {
  switch (functionId) {
    case 0x0000:
      return false;
    case 0x001e:
      renderer.stack.push(cloneState(renderer.state));
      renderer.context.save();
      break;
    case 0x0102:
      renderer.state.backgroundMode = view.getUint16(data, true);
      break;
    case 0x0106:
      renderer.state.fillMode = view.getUint16(data, true) === 1 ? 'evenodd' : 'nonzero';
      break;
    case 0x012d:
      renderer.selectObject(view.getUint16(data, true));
      break;
    case 0x012e:
      renderer.state.textAlign = view.getUint16(data, true);
      break;
    case 0x0127: {
      const restored = renderer.stack.pop();
      if (restored) {
        renderer.state = restored;
        renderer.context.restore();
      }
      break;
    }
    case 0x01f0:
      renderer.objects.delete(view.getUint16(data, true));
      break;
    case 0x0201:
      renderer.state.backgroundColor = colorRef(view, data);
      break;
    case 0x0209:
      renderer.state.textColor = colorRef(view, data);
      break;
    case 0x020b:
      renderer.state.windowOrigin = {
        x: view.getInt16(data + 2, true),
        y: view.getInt16(data, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 0x020c:
      renderer.state.windowExtent = {
        x: view.getInt16(data + 2, true),
        y: view.getInt16(data, true),
      };
      renderer.state.mappingEnabled = true;
      break;
    case 0x0214:
      renderer.state.current = {
        x: view.getInt16(data + 2, true),
        y: view.getInt16(data, true),
      };
      break;
    case 0x0213: {
      const end = { x: view.getInt16(data + 2, true), y: view.getInt16(data, true) };
      drawPoints(renderer, [renderer.state.current, end], false);
      renderer.state.current = end;
      break;
    }
    case 0x02fa:
      addWmfObject(renderer, {
        kind: 'pen',
        style: view.getUint16(data, true),
        width: Math.abs(view.getInt16(data + 2, true)) || 1,
        color: colorRef(view, data + 6),
        null: false,
      });
      break;
    case 0x02fc: {
      const style = view.getUint16(data, true);
      addWmfObject(renderer, {
        kind: 'brush',
        style,
        color: colorRef(view, data + 2),
        null: style === 1,
      });
      break;
    }
    case 0x02fb:
      addWmfObject(renderer, {
        kind: 'font',
        height: Math.abs(view.getInt16(data, true)) || 12,
        escapement: view.getInt16(data + 4, true),
        weight: view.getInt16(data + 8, true) || 400,
        italic: view.getUint8(data + 10) !== 0,
        underline: view.getUint8(data + 11) !== 0,
        strike: view.getUint8(data + 12) !== 0,
        family: readAnsi(view, data + 18, Math.max(0, dataLength - 18)) || 'sans-serif',
      });
      break;
    case 0x0324:
    case 0x0325: {
      const count = view.getUint16(data, true);
      if (2 + count * 4 > dataLength) break;
      const points: Point[] = [];
      for (let index = 0; index < count; index += 1) {
        points.push(wmfPoint(view, data + 2 + index * 4));
      }
      drawPoints(renderer, points, functionId === 0x0324);
      break;
    }
    case 0x0418: {
      const bottom = view.getInt16(data, true);
      const right = view.getInt16(data + 2, true);
      const top = view.getInt16(data + 4, true);
      const left = view.getInt16(data + 6, true);
      drawEllipse(renderer, { left, top, right, bottom });
      break;
    }
    case 0x041b: {
      const bottom = view.getInt16(data, true);
      const right = view.getInt16(data + 2, true);
      const top = view.getInt16(data + 4, true);
      const left = view.getInt16(data + 6, true);
      drawRectangle(renderer, { left, top, right, bottom });
      break;
    }
    case 0x061c: {
      const bottom = view.getInt16(data + 4, true);
      const right = view.getInt16(data + 6, true);
      const top = view.getInt16(data + 8, true);
      const left = view.getInt16(data + 10, true);
      drawRectangle(
        renderer,
        { left, top, right, bottom },
        {
          x: view.getInt16(data + 2, true),
          y: view.getInt16(data, true),
        },
      );
      break;
    }
    case 0x0521: {
      const count = view.getUint16(data, true);
      const padded = (count + 1) & ~1;
      if (2 + padded + 4 > dataLength) break;
      renderer.drawText(
        readAnsi(view, data + 2, count),
        view.getInt16(data + 2 + padded + 2, true),
        view.getInt16(data + 2 + padded, true),
      );
      break;
    }
    case 0x0a32: {
      const count = view.getUint16(data + 4, true);
      const options = view.getUint16(data + 6, true);
      const stringOffset = data + ((options & 6) !== 0 ? 16 : 8);
      if (stringOffset + count > data + dataLength) break;
      renderer.drawText(
        readAnsi(view, stringOffset, count),
        view.getInt16(data + 2, true),
        view.getInt16(data, true),
      );
      break;
    }
  }
  return true;
}

export async function renderWmfToDataUrl(
  buffer: ArrayBuffer,
  maxWidth?: number,
  maxHeight?: number,
  optionsOrScale?: MetafileRenderOptions | number,
): Promise<string | null> {
  try {
    const view = new DataView(buffer);
    const header = wmfHeader(view);
    if (!header) return null;
    const options = readOptions(maxWidth, maxHeight, optionsOrScale);
    const surface = createSurface(
      header.bounds.right - header.bounds.left,
      header.bounds.bottom - header.bounds.top,
      options,
    );
    if (!surface) return null;
    const renderer = new GdiRenderer(
      surface.context,
      header.bounds,
      surface.scaleX,
      surface.scaleY,
      options.fontFamilyMap,
    );
    renderer.state.mappingEnabled = true;
    let offset = header.recordsOffset;
    let records = 0;
    while (offset + 6 <= view.byteLength && records < options.maxRecords) {
      const words = view.getUint32(offset, true);
      const size = words * 2;
      if (words < 3 || !validRange(view, offset, size)) return null;
      const functionId = view.getUint16(offset + 4, true);
      records += 1;
      const keepGoing = replayWmfRecord(renderer, view, offset + 6, functionId, size - 6);
      offset += size;
      if (!keepGoing) break;
    }
    return await exportPng(surface.canvas);
  } catch {
    return null;
  }
}
