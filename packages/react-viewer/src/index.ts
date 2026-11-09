import './styles.css';

export { ReactPptxViewer } from './ReactPptxViewer';
export { parsePresentation, presentationParsingDefaults } from './parse';
export { initWasm, setWasmSource } from './wasm';
export type { WasmSource, WorkerWasmSource } from './wasm';
export { PptxViewerError } from './errors';
export { usePptxModel, usePptxPresentation, usePptxViewer } from './hooks';
export { usePptxViewerThumbnails } from './thumbnails';
export { createPptxThumbnailRenderer } from './thumbnail-renderer';
export { OFFICE_FONT_FALLBACKS, resolvePptxFontFamily } from './fonts';
export type * from './types';
export type { UsePptxPresentationState, UsePptxViewerResult } from './hooks';

export { extractPresentationBlocks } from './wasm';
export { flattenPresentationBlocks } from '@extend-ai/react-pptx-model';
export type {
  PptxSourceRef,
  PresentationBlock,
  PresentationWord,
  StructuredPresentation,
  ExtractPresentationBlocksOptions,
} from '@extend-ai/react-pptx-model';
export * from './block-geometry';
