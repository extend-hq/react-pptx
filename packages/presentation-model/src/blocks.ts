export interface PptxSourceRef {
  documentId: string;
  slideIndex: number;
  slideId: string;
  nodeId: string;
  nodePath: number[];
  partName?: string;
  sourceId?: string;
  story: 'slide' | 'notes';
  rowIndex?: number;
  columnIndex?: number;
  paragraphIndex?: number;
  start?: number;
  end?: number;
  offsetEncoding: 'utf-16';
}

export interface PresentationWord {
  object: 'word';
  id: string;
  blockId: string;
  content: string;
  sourceRefs: PptxSourceRef[];
}

export interface PresentationBlock {
  object: 'block';
  id: string;
  type: 'text' | 'text_box' | 'figure' | 'table' | 'table_cell' | 'group' | 'note';
  content: string;
  details: Record<string, unknown>;
  sourceRefs: PptxSourceRef[];
  parentBlockId?: string;
  children: PresentationBlock[];
  words?: PresentationWord[];
}

export interface StructuredPresentation {
  schemaVersion: 1;
  parserVersion: string;
  documentId: string;
  blocks: PresentationBlock[];
  geometryStatus: 'requires_layout';
}

export interface ExtractPresentationBlocksOptions {
  includeWords?: boolean;
  includeNotes?: boolean;
}

export function flattenPresentationBlocks(
  blocks: readonly PresentationBlock[],
): PresentationBlock[] {
  return blocks.flatMap((block) => [block, ...flattenPresentationBlocks(block.children)]);
}
