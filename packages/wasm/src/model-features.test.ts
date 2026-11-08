import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import type {
  AnimationNode,
  PresentationDocument,
  ShapeNode,
  TableNode,
} from '@extend-ai/react-pptx-model';

import initialize, { parse_presentation } from '../pkg/pptx_wasm.js';
import {
  buildModelFeaturesDeck,
  SLIDE_NUMBER_FIELD_ID,
} from '../../../scripts/model-features-deck.mjs';

async function parseModelFeaturesDeck(): Promise<PresentationDocument> {
  const workspaceRoot = [process.cwd(), resolve(process.cwd(), '../..')].find((directory) =>
    existsSync(resolve(directory, 'packages/wasm/pkg/pptx_wasm_bg.wasm')),
  );
  if (!workspaceRoot) throw new Error('Could not locate the react-pptx workspace root.');
  const wasm = await readFile(resolve(workspaceRoot, 'packages/wasm/pkg/pptx_wasm_bg.wasm'));
  await initialize({ module_or_path: wasm });
  return parse_presentation(buildModelFeaturesDeck()) as PresentationDocument;
}

function shapeNamed(document: PresentationDocument, slide: number, name: string): ShapeNode {
  const node = document.slides[slide]!.nodes.find(
    (candidate): candidate is ShapeNode => candidate.type === 'shape' && candidate.name === name,
  );
  if (!node) throw new Error(`Missing shape ${name} on slide ${slide}`);
  return node;
}

function flattenTiming(nodes: readonly AnimationNode[]): AnimationNode[] {
  return nodes.flatMap((node) => [node, ...flattenTiming(node.time?.children ?? [])]);
}

describe('presentation model features emitted by the Rust parser', () => {
  it('parses transitions, timing, math, fields, merges, actions and accessibility metadata', async () => {
    const document = await parseModelFeaturesDeck();
    expect(document.warnings).toEqual([]);
    expect(document.slides).toHaveLength(2);
    const [first, second] = document.slides as [
      PresentationDocument['slides'][number],
      PresentationDocument['slides'][number],
    ];

    // Slide transitions keep the effect, speed, advance settings and source.
    expect(first.transition).toMatchObject({
      effect: { kind: 'fade', throughBlack: true, parameters: { thruBlk: '1' } },
      speed: 'slow',
      advanceOnClick: true,
      advanceAfterMs: 3000,
    });
    expect(first.transition?.sourceXml).toContain('<p:transition');
    expect(second.transition).toMatchObject({
      effect: { kind: 'push', direction: 'l' },
      speed: 'med',
    });

    // The timing tree mirrors p:timing: the main sequence, the click effect
    // and the set behaviour that targets shape 2, plus the build list.
    const timing = first.timing!;
    expect(timing.builds).toEqual([
      expect.objectContaining({ kind: 'bldP', shapeId: '2', groupId: 0, build: 'whole' }),
    ]);
    const nodes = flattenTiming(timing.nodes ?? []);
    const root = nodes[0]!;
    expect(root.time).toMatchObject({ id: 1, nodeType: 'tmRoot', durationMs: 'indefinite', repeatCount: 1 });
    const sequence = nodes.find((node) => node.kind === 'seq')!;
    expect(sequence.sequence).toMatchObject({
      concurrent: true,
      nextAction: 'seek',
      nextConditions: [{ event: 'onNext', delayMs: 0, target: { kind: 'sldTgt' } }],
    });
    const effect = nodes.find((node) => node.time?.nodeType === 'clickEffect')!;
    expect(effect.time).toMatchObject({ presetId: 1, presetClass: 'entr', presetSubtype: 0 });
    const behaviour = nodes.find((node) => node.kind === 'set')!;
    expect(behaviour.behavior).toMatchObject({
      target: { kind: 'spTgt', shapeId: '2' },
      attributeNames: ['style.visibility'],
    });
    expect(behaviour.parameters.to).toEqual({ kind: 'string', value: 'visible' });
    expect(behaviour.time?.durationMs).toBe(1);
    // The build target resolves to the equation shape through its source id.
    const equation = shapeNamed(document, 0, 'Equation');
    expect(equation.sourceId).toBe('2');
    expect(equation.sourcePart).toBe('ppt/slides/slide1.xml');

    // Office Math runs keep the OMML tree and a MathML rendering.
    const mathRun = equation.paragraphs![0]!.runs.find((run) => run.math)!;
    expect(mathRun.text).toBe('E=mc2');
    expect(mathRun.math).toMatchObject({ display: true, text: 'E=mc2' });
    expect(mathRun.math?.mathMl).toMatch(/^<math[\s>]/);
    expect(mathRun.math?.mathMl).toContain('<msup>');
    expect(mathRun.math?.root).toMatchObject({ name: 'oMathPara', children: [{ name: 'oMath' }] });
    expect(mathRun.math?.unsupported ?? []).toEqual([]);

    // Text fields keep their id, type and cached text.
    const fieldRun = shapeNamed(document, 0, 'Slide number').paragraphs![0]!.runs[0]!;
    expect(fieldRun.text).toBe('1');
    expect(fieldRun.field).toMatchObject({
      id: SLIDE_NUMBER_FIELD_ID,
      type: 'slidenum',
      hasCachedText: true,
      sourcePart: 'ppt/slides/slide1.xml',
    });

    // Click actions distinguish presentation jumps from slide links, and
    // non-visual properties surface as accessibility metadata.
    const next = shapeNamed(document, 0, 'Next button');
    expect(next.clickAction).toMatchObject({ action: 'ppaction://hlinkshowjump?jump=nextslide' });
    expect(next.accessibilityTitle).toBe('Next');
    expect(next.description).toBe('Advances to the next slide');
    expect(next.altText).toBe('Advances to the next slide');
    const link = shapeNamed(document, 0, 'Summary link');
    expect(link.clickAction).toMatchObject({
      external: false,
      relationshipId: 'rId2',
      target: 'ppt/slides/slide2.xml',
      action: 'ppaction://hlinksldjump',
    });
    expect(second.sourcePart).toBe('ppt/slides/slide2.xml');
    expect(shapeNamed(document, 1, 'Hidden note').hidden).toBe(true);
    expect(shapeNamed(document, 1, 'Decorative bar').decorative).toBe(true);
    expect(equation.editing).toMatchObject({ kind: 'shape', textBox: true });

    // Merged table cells point back at their origin cell.
    const table = first.nodes.find((node): node is TableNode => node.type === 'table')!;
    expect(table.editing).toMatchObject({ kind: 'graphicFrame', locks: { noGrouping: true } });
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0]![0]).toMatchObject({ colSpan: 2 });
    expect(table.rows[0]![1]!.merge).toEqual({ horizontal: true, vertical: false, origin: { row: 0, column: 0 } });
    expect(table.rows[1]![0]).toMatchObject({ rowSpan: 2 });
    expect(table.rows[2]![0]!.merge).toEqual({ horizontal: false, vertical: true, origin: { row: 1, column: 0 } });
    expect(table.rows[2]![1]!.merge).toBeUndefined();
  });
});
