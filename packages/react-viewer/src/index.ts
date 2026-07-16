import './styles.css';

export { ReactPptxViewer } from './ReactPptxViewer';
export { parsePresentation, presentationParsingDefaults } from './parse';
export { initWasm, setWasmSource } from './wasm';
export type { WasmSource, WorkerWasmSource } from './wasm';
export { PptxViewerError } from './errors';
export { usePptxModel, usePptxPresentation, usePptxViewer } from './hooks';
export { usePptxViewerThumbnails } from './thumbnails';
export { OFFICE_FONT_FALLBACKS, resolvePptxFontFamily } from './fonts';
export type * from './types';
export type { UsePptxPresentationState, UsePptxViewerResult } from './hooks';
