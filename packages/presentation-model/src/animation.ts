import type { ColorValue } from './index';

/** Finite quantity or indefinite; durations are milliseconds and percentages are fractions. */
export type AnimationNumber = number | 'indefinite';
export type AnimationValue =
  | { kind: 'string'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'integer' | 'number'; value: number }
  | { kind: 'color'; value: ColorValue }
  | { kind: 'point'; value: AnimationPoint }
  | { kind: 'rgbDelta' | 'hslDelta'; value: [number, number, number] };

export interface SlideTransition {
  effect?: TransitionEffect;
  speed: string;
  durationMs?: number;
  advanceOnClick: boolean;
  advanceAfterMs?: number;
  sound?: TransitionSound;
  sourceXml: string;
}

export interface TransitionEffect {
  kind: string;
  namespace: string;
  direction?: string;
  orientation?: string;
  throughBlack?: boolean;
  spokes?: number;
  parameters: Record<string, string>;
}

export interface TransitionSound {
  action: string;
  loopSound: boolean;
  source?: AnimationSound;
}

export interface AnimationSound {
  relationshipId?: string;
  target?: string;
  assetId?: string;
  external: boolean;
  name?: string;
  builtin?: boolean;
}

export interface SlideTiming {
  nodes?: AnimationNode[];
  builds?: AnimationBuild[];
  unsupported?: string[];
  sourceXml: string;
}

export interface AnimationNode {
  kind: string;
  time?: AnimationTime;
  behavior?: AnimationBehavior;
  sequence?: AnimationSequence;
  media?: AnimationMedia;
  parameters: AnimationParameters;
}

export interface AnimationTime {
  id?: number;
  presetId?: number;
  presetClass?: string;
  presetSubtype?: number;
  durationMs?: AnimationNumber;
  repeatCount: AnimationNumber;
  repeatDurationMs?: AnimationNumber;
  speed: number;
  acceleration: number;
  deceleration: number;
  autoReverse: boolean;
  restart?: string;
  fill?: string;
  syncBehavior?: string;
  timeFilter?: string;
  eventFilter?: string;
  display?: boolean;
  masterRelation?: string;
  buildLevel?: number;
  groupId?: number;
  afterEffect?: boolean;
  nodeType?: string;
  placeholder?: boolean;
  startConditions?: AnimationCondition[];
  endConditions?: AnimationCondition[];
  endSync?: AnimationCondition;
  iterate?: AnimationIteration;
  children?: AnimationNode[];
  subNodes?: AnimationNode[];
}

export interface AnimationCondition {
  event?: string;
  delayMs?: AnimationNumber;
  target?: AnimationTarget;
  timeNodeId?: number;
  runtimeNode?: string;
}

export interface AnimationIteration {
  kind: string;
  backwards: boolean;
  intervalMs?: AnimationNumber;
  intervalFraction?: number;
}

export interface AnimationTarget {
  kind: string;
  /** Identifier scoped to the owning part, matching object source identity. */
  shapeId?: string;
  subShapeId?: string;
  background?: boolean;
  text?: AnimationTextRange;
  chart?: AnimationChartTarget;
  diagram?: AnimationDiagramTarget;
  sound?: AnimationSound;
}

export interface AnimationTextRange {
  kind: string;
  start?: number;
  end?: number;
}

export interface AnimationChartTarget {
  kind: string;
  level?: number;
  seriesIndex?: number;
  categoryIndex?: number;
}

export interface AnimationDiagramTarget {
  id?: string;
  buildStep?: string;
}

export interface AnimationBehavior {
  target?: AnimationTarget;
  attributeNames?: string[];
  additive?: string;
  accumulate?: string;
  transformType?: string;
  from?: string;
  to?: string;
  by?: string;
  runtimeContext?: string;
  overrideMode?: string;
}

export interface AnimationSequence {
  concurrent?: boolean;
  previousAction?: string;
  nextAction?: string;
  previousConditions?: AnimationCondition[];
  nextConditions?: AnimationCondition[];
}

export interface AnimationMedia {
  target?: AnimationTarget;
  volume: number;
  mute: boolean;
  slideCount: number;
  showWhenStopped: boolean;
  narration?: boolean;
  fullScreen?: boolean;
}

export interface AnimationParameters {
  from?: AnimationValue;
  to?: AnimationValue;
  by?: AnimationValue;
  calculationMode?: string;
  valueType?: string;
  keyframes?: AnimationKeyframe[];
  colorSpace?: string;
  direction?: string;
  transition?: string;
  filter?: string;
  properties?: string;
  progress?: AnimationValue;
  origin?: string;
  path?: string;
  pathEditMode?: string;
  rotationCenter?: AnimationPoint;
  rotationDegrees?: number;
  pointTypes?: string;
  zoomContents?: boolean;
  commandType?: string;
  command?: string;
}

export interface AnimationPoint {
  x: number;
  y: number;
}

export interface AnimationKeyframe {
  time: AnimationNumber;
  formula: string;
  value?: AnimationValue;
}

export interface AnimationBuild {
  kind: string;
  /** Identifier scoped to the owning part, matching object source identity. */
  shapeId?: string;
  groupId?: number;
  uiExpand: boolean;
  build?: string;
  level?: number;
  animateBackground?: boolean;
  autoUpdateBackground?: boolean;
  reverse?: boolean;
  advanceAfterMs?: AnimationNumber;
  asOne?: boolean;
  graphicBuild?: AnimationGraphicBuild;
  templates?: AnimationTemplate[];
}

export interface AnimationGraphicBuild {
  kind: string;
  build?: string;
  reverse?: boolean;
  animateBackground?: boolean;
}

export interface AnimationTemplate {
  level: number;
  nodes?: AnimationNode[];
}

