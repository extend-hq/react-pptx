import * as React from 'react';
import { parsePresentation, type ParsedPresentation } from '@extend-ai/react-pptx';
import { editPptx, type ObjectEdit } from './pptx-editing';

export function usePptxEditing(
  opened: { source: string | File },
  onChange: (parsed: ParsedPresentation) => void,
) {
  const bytes = React.useRef<Uint8Array | null>(null);
  const history = React.useRef<{ undo: Uint8Array[]; redo: Uint8Array[] }>({ undo: [], redo: [] });
  const generation = React.useRef(0);
  const working = React.useRef(false);
  const [ready, setReady] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [revision, setRevision] = React.useState(0);

  React.useEffect(() => {
    const current = ++generation.current;
    const abort = new AbortController();
    bytes.current = null;
    history.current = { undo: [], redo: [] };
    working.current = false;
    setReady(false);
    setBusy(false);
    setError(null);
    setRevision(0);
    void (async () => {
      try {
        let buffer: ArrayBuffer;
        if (typeof opened.source === 'string') {
          const response = await fetch(opened.source, { signal: abort.signal });
          if (!response.ok) throw new Error('The presentation could not be loaded for editing.');
          buffer = await response.arrayBuffer();
        } else buffer = await opened.source.arrayBuffer();
        if (generation.current !== current || abort.signal.aborted) return;
        bytes.current = new Uint8Array(buffer);
        setReady(true);
      } catch (cause) {
        if (!abort.signal.aborted && generation.current === current) {
          setError(cause instanceof Error ? cause.message : 'Editing is unavailable.');
        }
      }
    })();
    return () => {
      abort.abort();
      generation.current++;
    };
  }, [opened]);

  const update = React.useCallback(
    async (edit: ObjectEdit | 'undo' | 'redo') => {
      const previous = bytes.current;
      if (!previous || working.current) return;
      const current = generation.current;
      working.current = true;
      setBusy(true);
      setError(null);
      try {
        const next =
          typeof edit === 'string' ? history.current[edit].at(-1) : await editPptx(previous, edit);
        if (!next) return;
        const parsed = await parsePresentation(next);
        if (current !== generation.current) return;
        if (typeof edit === 'string') {
          history.current[edit].pop();
          history.current[edit === 'undo' ? 'redo' : 'undo'].push(previous);
        } else {
          history.current.undo.push(previous);
          history.current.redo = [];
        }
        while (
          history.current.undo.length > 20 ||
          (history.current.undo.length > 1 &&
            history.current.undo.reduce((sum, entry) => sum + entry.byteLength, 0) >
              64 * 1024 * 1024)
        ) {
          history.current.undo.shift();
        }
        bytes.current = next;
        setRevision((value) => value + 1);
        onChange(parsed);
      } catch (cause) {
        if (current === generation.current)
          setError(cause instanceof Error ? cause.message : 'The change could not be saved.');
      } finally {
        if (current === generation.current) {
          working.current = false;
          setBusy(false);
        }
      }
    },
    [onChange],
  );

  return {
    ready,
    busy,
    error,
    revision,
    canUndo: history.current.undo.length > 0,
    canRedo: history.current.redo.length > 0,
    update,
    download: (name: string) => {
      if (!bytes.current || working.current) return;
      const href = URL.createObjectURL(
        new Blob([new Uint8Array(bytes.current)], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        }),
      );
      const link = document.createElement('a');
      link.href = href;
      link.download = name;
      link.click();
      setTimeout(() => URL.revokeObjectURL(href), 1000);
    },
  };
}
