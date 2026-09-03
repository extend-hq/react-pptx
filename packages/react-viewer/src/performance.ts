import type { PptxPerformanceMeasurement, PptxPerformancePhase } from './types';

export function performanceNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

export function performanceMeasurement(
  phase: PptxPerformancePhase,
  startedAt: number,
  slideIndex?: number,
): PptxPerformanceMeasurement {
  return {
    durationMs: Math.max(0, performanceNow() - startedAt),
    phase,
    ...(slideIndex === undefined ? {} : { slideIndex }),
  };
}
