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

type PathCommand =
  | { kind: 'move' | 'line'; point: Point }
  | { kind: 'bezier'; first: Point; second: Point; end: Point }
  | { kind: 'close' };

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

type GdiObject = Pen | Brush | Font | { kind: 'unsupported' };

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

function multiplyMatrices(left: Matrix, right: Matrix): Matrix {
  return {
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  };
}

function readMatrix(view: DataView, offset: number): Matrix {
  return {
    a: view.getFloat32(offset, true),
    b: view.getFloat32(offset + 4, true),
    c: view.getFloat32(offset + 8, true),
    d: view.getFloat32(offset + 12, true),
    e: view.getFloat32(offset + 16, true),
    f: view.getFloat32(offset + 20, true),
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
    fillMode: 'evenodd',
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
  path: PathCommand[] = [];
  pathOpen = false;

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
    const origin = this.point(0, 0);
    const end = this.point(value, 0);
    return Math.hypot(end.x - origin.x, end.y - origin.y);
  }

  lengthY(value: number): number {
    const origin = this.point(0, 0);
    const end = this.point(0, value);
    return Math.hypot(end.x - origin.x, end.y - origin.y);
  }

  applyPen(): boolean {
    const pen = this.state.pen;
    if (pen.null) return false;
    this.context.strokeStyle = pen.color;
    this.context.lineWidth = Math.max(1, this.lengthX(pen.width || 1));
    const cap = pen.style & 0x0f00;
    const join = pen.style & 0xf000;
    this.context.lineCap = cap === 0x0100 ? 'square' : cap === 0x0200 ? 'butt' : 'round';
    this.context.lineJoin = join === 0x1000 ? 'bevel' : join === 0x2000 ? 'miter' : 'round';
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

  paintPath(fill: boolean, stroke = true): void {
    if (fill && this.applyBrush()) this.context.fill(this.state.fillMode);
    if (stroke && this.applyPen()) this.context.stroke();
  }

  replayPath(commands: PathCommand[], closeFigures = false): void {
    this.context.beginPath();
    for (const command of commands) {
      if (command.kind === 'move') {
        if (closeFigures) this.context.closePath();
        this.context.moveTo(command.point.x, command.point.y);
      } else if (command.kind === 'line') this.context.lineTo(command.point.x, command.point.y);
      else if (command.kind === 'bezier') {
        this.context.bezierCurveTo(
          command.first.x,
          command.first.y,
          command.second.x,
          command.second.y,
          command.end.x,
          command.end.y,
        );
      } else this.context.closePath();
    }
    if (closeFigures) this.context.closePath();
  }

  drawPath(commands: PathCommand[], fill: boolean): void {
    const transformed = commands.map((command): PathCommand => {
      if (command.kind === 'close') return command;
      if (command.kind === 'bezier') {
        return {
          kind: 'bezier',
          first: this.point(command.first.x, command.first.y),
          second: this.point(command.second.x, command.second.y),
          end: this.point(command.end.x, command.end.y),
        };
      }
      return { kind: command.kind, point: this.point(command.point.x, command.point.y) };
    });
    if (this.pathOpen) {
      for (const command of transformed) this.path.push(command);
    } else {
      this.replayPath(transformed);
      this.paintPath(fill);
    }
  }

  restore(level: number): void {
    const index = level < 0 ? this.stack.length + level : level - 1;
    const restored = this.stack[index];
    if (!restored) return;
    while (this.stack.length > index) {
      this.stack.pop();
      this.context.restore();
    }
    this.state = restored;
  }

  selectObject(index: number): void {
    const object = index >= 0x80000000 ? stockObject(index & 0x7fffffff) : this.objects.get(index);
    if (!object) return;
    if (object.kind === 'pen') this.state.pen = { ...object };
    else if (object.kind === 'brush') this.state.brush = { ...object };
    else if (object.kind === 'font') this.state.font = { ...object };
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
  if (index >= 10 && index <= 16) return { ...DEFAULT_FONT };
  return undefined;
}

function validRange(view: DataView, offset: number, length: number): boolean {
  return offset >= 0 && length >= 0 && offset <= view.byteLength - length;
}

function rectanglePath(bounds: Bounds, radius?: Point): PathCommand[] {
  const left = Math.min(bounds.left, bounds.right);
  const right = Math.max(bounds.left, bounds.right);
  const top = Math.min(bounds.top, bounds.bottom);
  const bottom = Math.max(bounds.top, bounds.bottom);
  const rx = Math.min(Math.abs(radius?.x ?? 0), right - left) / 2;
  const ry = Math.min(Math.abs(radius?.y ?? 0), bottom - top) / 2;
  if (!rx || !ry) {
    return [
      { kind: 'move', point: { x: left, y: top } },
      { kind: 'line', point: { x: right, y: top } },
      { kind: 'line', point: { x: right, y: bottom } },
      { kind: 'line', point: { x: left, y: bottom } },
      { kind: 'close' },
    ];
  }
  const k = 0.5522847498307936;
  return [
    { kind: 'move', point: { x: left + rx, y: top } },
    { kind: 'line', point: { x: right - rx, y: top } },
    {
      kind: 'bezier',
      first: { x: right - rx * (1 - k), y: top },
      second: { x: right, y: top + ry * (1 - k) },
      end: { x: right, y: top + ry },
    },
    { kind: 'line', point: { x: right, y: bottom - ry } },
    {
      kind: 'bezier',
      first: { x: right, y: bottom - ry * (1 - k) },
      second: { x: right - rx * (1 - k), y: bottom },
      end: { x: right - rx, y: bottom },
    },
    { kind: 'line', point: { x: left + rx, y: bottom } },
    {
      kind: 'bezier',
      first: { x: left + rx * (1 - k), y: bottom },
      second: { x: left, y: bottom - ry * (1 - k) },
      end: { x: left, y: bottom - ry },
    },
    { kind: 'line', point: { x: left, y: top + ry } },
    {
      kind: 'bezier',
      first: { x: left, y: top + ry * (1 - k) },
      second: { x: left + rx * (1 - k), y: top },
      end: { x: left + rx, y: top },
    },
    { kind: 'close' },
  ];
}

function drawRectangle(renderer: GdiRenderer, bounds: Bounds, radius?: Point): void {
  renderer.drawPath(rectanglePath(bounds, radius), true);
}

function drawEllipse(renderer: GdiRenderer, bounds: Bounds): void {
  drawRectangle(renderer, bounds, {
    x: bounds.right - bounds.left,
    y: bounds.bottom - bounds.top,
  });
}

function drawBrushBlit(renderer: GdiRenderer, bounds: Bounds, operation: number): void {
  const raster = (operation >>> 16) & 0xff;
  if (raster === 0xf0) {
    if (!renderer.applyBrush()) return;
  } else if (raster === 0x00) renderer.context.fillStyle = '#000000';
  else if (raster === 0xff) renderer.context.fillStyle = '#ffffff';
  else return;
  const corners = [
    renderer.point(bounds.left, bounds.top),
    renderer.point(bounds.right, bounds.top),
    renderer.point(bounds.right, bounds.bottom),
    renderer.point(bounds.left, bounds.bottom),
  ];
  renderer.replayPath(pointCommands(corners, true));
  renderer.context.fill();
}

interface RgbBitmap {
  canvas: RenderCanvas;
  height: number;
  bottomUp: boolean;
}

function readRgbBitmap(
  view: DataView,
  infoOffset: number,
  infoLength: number,
  bitsOffset: number,
  bitsLength: number,
): RgbBitmap | undefined {
  if (
    infoLength < 40 ||
    !validRange(view, infoOffset, infoLength) ||
    !validRange(view, bitsOffset, bitsLength)
  ) {
    return undefined;
  }
  const headerSize = view.getUint32(infoOffset, true);
  const width = view.getInt32(infoOffset + 4, true);
  const signedHeight = view.getInt32(infoOffset + 8, true);
  const height = Math.abs(signedHeight);
  const planes = view.getUint16(infoOffset + 12, true);
  const depth = view.getUint16(infoOffset + 14, true);
  const compression = view.getUint32(infoOffset + 16, true);
  if (
    headerSize < 40 ||
    headerSize > infoLength ||
    width < 1 ||
    height < 1 ||
    width > DEFAULT_MAX_DIMENSION ||
    height > DEFAULT_MAX_DIMENSION ||
    planes !== 1 ||
    (depth !== 24 && depth !== 32) ||
    compression !== 0
  ) {
    return undefined;
  }
  const stride = Math.ceil((width * depth) / 32) * 4;
  if (stride * height > bitsLength) return undefined;
  const surface = createSurface(width, height, readOptions(undefined, undefined, undefined));
  if (!surface) return undefined;
  const pixels = surface.context.createImageData(width, height);
  for (let y = 0; y < height; y += 1) {
    const row = signedHeight > 0 ? height - y - 1 : y;
    for (let x = 0; x < width; x += 1) {
      const source = bitsOffset + row * stride + (x * depth) / 8;
      const target = (y * width + x) * 4;
      pixels.data[target] = view.getUint8(source + 2);
      pixels.data[target + 1] = view.getUint8(source + 1);
      pixels.data[target + 2] = view.getUint8(source);
      pixels.data[target + 3] = 255;
    }
  }
  surface.context.putImageData(pixels, 0, 0);
  return { canvas: surface.canvas, height, bottomUp: signedHeight > 0 };
}

function drawRgbBitmap(
  renderer: GdiRenderer,
  bitmap: RgbBitmap,
  source: Bounds,
  destination: Bounds,
): void {
  const width = source.right - source.left;
  const height = source.bottom - source.top;
  if (!width || !height) return;
  const x = width < 0 ? destination.right : destination.left;
  const y = height < 0 ? destination.bottom : destination.top;
  const origin = renderer.point(x, y);
  const horizontal = renderer.point(width < 0 ? destination.left : destination.right, y);
  const vertical = renderer.point(x, height < 0 ? destination.top : destination.bottom);
  renderer.context.save();
  renderer.context.transform(
    horizontal.x - origin.x,
    horizontal.y - origin.y,
    vertical.x - origin.x,
    vertical.y - origin.y,
    origin.x,
    origin.y,
  );
  renderer.context.drawImage(
    bitmap.canvas,
    Math.min(source.left, source.right),
    bitmap.bottomUp
      ? bitmap.height - Math.max(source.top, source.bottom)
      : Math.min(source.top, source.bottom),
    Math.abs(width),
    Math.abs(height),
    0,
    0,
    1,
    1,
  );
  renderer.context.restore();
}

function pointCommands(points: Point[], close: boolean): PathCommand[] {
  const commands: PathCommand[] = [];
  for (const [index, point] of points.entries()) {
    commands.push({ kind: index === 0 ? 'move' : 'line', point });
  }
  if (close && points.length > 0) commands.push({ kind: 'close' });
  return commands;
}

function drawPoints(renderer: GdiRenderer, points: Point[], close: boolean): void {
  if (points.length > 0) renderer.drawPath(pointCommands(points, close), close);
}

function drawConnectedPoints(
  renderer: GdiRenderer,
  points: Point[],
  bezier: boolean,
  continues: boolean,
): void {
  const first = continues ? renderer.state.current : points[0];
  const offset = continues ? 0 : 1;
  if (!first || (bezier && (points.length - offset) % 3 !== 0)) return;
  const commands: PathCommand[] = [];
  const previous = renderer.path.at(-1);
  const previousEnd =
    previous?.kind === 'bezier'
      ? previous.end
      : previous?.kind === 'move' || previous?.kind === 'line'
        ? previous.point
        : undefined;
  const start = renderer.point(first.x, first.y);
  if (
    !continues ||
    !renderer.pathOpen ||
    previousEnd?.x !== start.x ||
    previousEnd?.y !== start.y
  ) {
    commands.push({ kind: 'move', point: first });
  }
  for (let index = offset; index < points.length; index += bezier ? 3 : 1) {
    if (bezier) {
      commands.push({
        kind: 'bezier',
        first: points[index]!,
        second: points[index + 1]!,
        end: points[index + 2]!,
      });
    } else commands.push({ kind: 'line', point: points[index]! });
  }
  renderer.drawPath(commands, false);
  if (continues && points.length > 0) renderer.state.current = points[points.length - 1]!;
}

function readPoints(view: DataView, offset: number, count: number, short: boolean): Point[] {
  const stride = short ? 4 : 8;
  if (count > 1_000_000 || !validRange(view, offset, count * stride)) {
    throw new RangeError('Invalid metafile point array');
  }
  const points: Point[] = [];
  for (let index = 0; index < count; index += 1) {
    const cursor = offset + index * stride;
    points.push({
      x: short ? view.getInt16(cursor, true) : view.getInt32(cursor, true),
      y: short ? view.getInt16(cursor + 2, true) : view.getInt32(cursor + 4, true),
    });
  }
  return points;
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
      if (renderer.pathOpen) {
        renderer.drawPath([{ kind: 'move', point: renderer.state.current }], false);
      }
      break;
    case 30: {
      const bounds = readBounds32(view, data);
      renderer.replayPath(
        rectanglePath(bounds).map((command) =>
          command.kind === 'move' || command.kind === 'line'
            ? { kind: command.kind, point: renderer.point(command.point.x, command.point.y) }
            : command,
        ),
      );
      renderer.context.clip();
      break;
    }
    case 33:
      renderer.stack.push(cloneState(renderer.state));
      renderer.context.save();
      break;
    case 34:
      renderer.restore(view.getInt32(data, true));
      break;
    case 35:
      renderer.state.world = readMatrix(view, data);
      break;
    case 36: {
      const matrix = readMatrix(view, data);
      const mode = view.getUint32(data + 24, true);
      if (mode === 1) renderer.state.world = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
      // XFORM uses row vectors; Canvas coordinates use the transposed matrix.
      else if (mode === 2) renderer.state.world = multiplyMatrices(renderer.state.world, matrix);
      else if (mode === 3) renderer.state.world = multiplyMatrices(matrix, renderer.state.world);
      else if (mode === 4) renderer.state.world = matrix;
      break;
    }
    case 37:
      renderer.selectObject(view.getUint32(data, true));
      break;
    case 38: {
      const style = view.getUint32(data + 4, true);
      renderer.objects.set(view.getUint32(data, true), {
        kind: 'pen',
        style,
        width: Math.abs(view.getInt32(data + 8, true)) || 1,
        color: colorRef(view, data + 16),
        null: (style & 0x0f) === 5,
      });
      break;
    }
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
      drawConnectedPoints(renderer, [end], false, true);
      break;
    }
    case 59:
      renderer.path = [];
      renderer.pathOpen = true;
      break;
    case 60:
      renderer.pathOpen = false;
      break;
    case 61:
      if (renderer.pathOpen) renderer.path.push({ kind: 'close' });
      break;
    case 62:
    case 63:
    case 64:
      if (!renderer.pathOpen) {
        renderer.replayPath(renderer.path, type !== 64);
        renderer.paintPath(type !== 64, type !== 62);
        renderer.path = [];
      }
      break;
    case 68:
      renderer.path = [];
      renderer.pathOpen = false;
      break;
    case 76: {
      const x = view.getInt32(data + 16, true);
      const y = view.getInt32(data + 20, true);
      drawBrushBlit(
        renderer,
        {
          left: x,
          top: y,
          right: x + view.getInt32(data + 24, true),
          bottom: y + view.getInt32(data + 28, true),
        },
        view.getUint32(data + 32, true),
      );
      break;
    }
    case 81: {
      const operation = view.getUint32(data + 60, true);
      const usage = view.getUint32(data + 56, true);
      if (operation !== 0x00cc0020 || usage !== 0) break;
      const bitmap = readRgbBitmap(
        view,
        recordOffset + view.getUint32(data + 40, true),
        view.getUint32(data + 44, true),
        recordOffset + view.getUint32(data + 48, true),
        view.getUint32(data + 52, true),
      );
      if (!bitmap) break;
      const x = view.getInt32(data + 16, true);
      const y = view.getInt32(data + 20, true);
      const sourceX = view.getInt32(data + 24, true);
      const sourceY = view.getInt32(data + 28, true);
      drawRgbBitmap(
        renderer,
        bitmap,
        {
          left: sourceX,
          top: sourceY,
          right: sourceX + view.getInt32(data + 32, true),
          bottom: sourceY + view.getInt32(data + 36, true),
        },
        {
          left: x,
          top: y,
          right: x + view.getInt32(data + 64, true),
          bottom: y + view.getInt32(data + 68, true),
        },
      );
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
    case 2:
    case 3:
    case 4:
    case 5:
    case 6:
    case 85:
    case 86:
    case 87:
    case 88:
    case 89: {
      const count = view.getUint32(data + 16, true);
      const short = type >= 85;
      const points = readPoints(view, data + 20, count, short);
      const kind = short ? type - 83 : type;
      if (kind === 3) drawPoints(renderer, points, true);
      else drawConnectedPoints(renderer, points, kind === 2 || kind === 5, kind >= 5);
      break;
    }
    case 7:
    case 8:
    case 90:
    case 91: {
      const polygons = view.getUint32(data + 16, true);
      const count = view.getUint32(data + 20, true);
      if (polygons > count || !validRange(view, data + 24, polygons * 4)) {
        throw new RangeError('Invalid metafile polygon array');
      }
      const points = readPoints(view, data + 24 + polygons * 4, count, type >= 90);
      const close = type === 8 || type === 91;
      const commands: PathCommand[] = [];
      let pointOffset = 0;
      for (let index = 0; index < polygons; index += 1) {
        const polygonCount = view.getUint32(data + 24 + index * 4, true);
        if (polygonCount > count - pointOffset) {
          throw new RangeError('Invalid metafile polygon size');
        }
        for (const command of pointCommands(
          points.slice(pointOffset, pointOffset + polygonCount),
          close,
        )) {
          commands.push(command);
        }
        pointOffset += polygonCount;
      }
      if (pointOffset !== count) throw new RangeError('Invalid metafile polygon count');
      renderer.drawPath(commands, close);
      break;
    }
    case 95: {
      const style = view.getUint32(data + 20, true);
      renderer.objects.set(view.getUint32(data, true), {
        kind: 'pen',
        style,
        width: view.getUint32(data + 24, true) || 1,
        color: colorRef(view, data + 32),
        null: (style & 0x0f) === 5 || view.getUint32(data + 28, true) === 1,
      });
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
      const keepGoing = replayEmfRecord(
        renderer,
        new DataView(buffer, offset, size),
        0,
        type,
        size,
      );
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
    case 0x00f7:
    case 0x0142:
      addWmfObject(renderer, { kind: 'unsupported' });
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
    case 0x0127:
      renderer.restore(view.getInt16(data, true));
      break;
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
    case 0x02fa: {
      const style = view.getUint16(data, true);
      addWmfObject(renderer, {
        kind: 'pen',
        style,
        width: Math.abs(view.getInt16(data + 2, true)) || 1,
        color: colorRef(view, data + 6),
        null: (style & 0x0f) === 5,
      });
      break;
    }
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
    case 0x061d: {
      const x = view.getInt16(data + 10, true);
      const y = view.getInt16(data + 8, true);
      drawBrushBlit(
        renderer,
        {
          left: x,
          top: y,
          right: x + view.getInt16(data + 6, true),
          bottom: y + view.getInt16(data + 4, true),
        },
        view.getUint32(data, true),
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
      {
        left: 0,
        top: 0,
        right: header.bounds.right - header.bounds.left,
        bottom: header.bounds.bottom - header.bounds.top,
      },
      surface.scaleX,
      surface.scaleY,
      options.fontFamilyMap,
    );
    renderer.state.windowOrigin = { x: header.bounds.left, y: header.bounds.top };
    renderer.state.mappingEnabled = true;
    let offset = header.recordsOffset;
    let records = 0;
    while (offset + 6 <= view.byteLength && records < options.maxRecords) {
      const words = view.getUint32(offset, true);
      const size = words * 2;
      if (words < 3 || !validRange(view, offset, size)) return null;
      const functionId = view.getUint16(offset + 4, true);
      records += 1;
      const keepGoing = replayWmfRecord(
        renderer,
        new DataView(buffer, offset, size),
        6,
        functionId,
        size - 6,
      );
      offset += size;
      if (!keepGoing) break;
    }
    return await exportPng(surface.canvas);
  } catch {
    return null;
  }
}
