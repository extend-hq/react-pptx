import { useCallback, useEffect, useState, type RefCallback } from 'react';
import { toPptxViewerError, type PptxViewerError } from './errors';
import { parsePresentation } from './parse';
import type {
  ParsePresentationOptions,
  ParsedPresentation,
  PptxViewerController,
  PresentationSource,
} from './types';

export interface UsePptxPresentationState {
  presentation: ParsedPresentation | null;
  document: ParsedPresentation['document'] | null;
  isLoading: boolean;
  error: PptxViewerError | null;
}

const EMPTY_PARSE_OPTIONS: ParsePresentationOptions = Object.freeze({});

/** Parse a presentation independently from the viewer, mirroring `useDocxModel`. */
export function usePptxPresentation(
  source?: PresentationSource | null,
  options: ParsePresentationOptions = EMPTY_PARSE_OPTIONS,
): UsePptxPresentationState {
  const [state, setState] = useState<UsePptxPresentationState>({
    presentation: null,
    document: null,
    isLoading: Boolean(source),
    error: null,
  });

  useEffect(() => {
    if (!source) {
      setState({ presentation: null, document: null, isLoading: false, error: null });
      return;
    }
    const abort = new AbortController();
    setState((current) => ({ ...current, isLoading: true, error: null }));
    void parsePresentation(source, { ...options, signal: abort.signal })
      .then((presentation) => {
        if (!abort.signal.aborted) {
          setState({
            presentation,
            document: presentation.document,
            isLoading: false,
            error: null,
          });
        }
      })
      .catch((reason: unknown) => {
        if (!abort.signal.aborted) {
          setState({
            presentation: null,
            document: null,
            isLoading: false,
            error: toPptxViewerError(reason, 'parse-failed', 'Could not open the presentation.'),
          });
        }
      });
    return () => abort.abort();
  }, [source, options.fetchInit, options.formatHint, options.maxInputBytes]);

  return state;
}

/** Backwards-friendly model-oriented alias. */
export const usePptxModel = usePptxPresentation;

export interface UsePptxViewerResult {
  controller: PptxViewerController | null;
  ref: RefCallback<PptxViewerController>;
}

/** Reactive access to the imperative viewer controller. */
export function usePptxViewer(): UsePptxViewerResult {
  const [controller, setController] = useState<PptxViewerController | null>(null);
  const ref = useCallback<RefCallback<PptxViewerController>>((value) => setController(value), []);
  return { controller, ref };
}
