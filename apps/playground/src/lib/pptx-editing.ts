import { strFromU8, strToU8, unzip, zip } from 'fflate';
import type { ObjectLocks, Rect, SlideNode } from '@extend-ai/react-pptx-model';

const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION = 'http://schemas.openxmlformats.org/presentationml/2006/main';

export type ObjectEdit = {
  node: SlideNode;
  bounds?: Rect;
  paragraph?: { index: number; text: string; row?: number; column?: number };
};

function children(element: Element, namespace: string, name: string) {
  return Array.from(element.children).filter(
    (child) => child.namespaceURI === namespace && child.localName === name,
  );
}

function child(element: Element, namespace: string, name: string) {
  return children(element, namespace, name)[0];
}

function drawingElement(parent: Element, name: string) {
  return parent.ownerDocument.createElementNS(DRAWING, `a:${name}`);
}

function replaceParagraphText(paragraph: Element, text: string) {
  const runs = Array.from(paragraph.children).filter(
    (element) => element.namespaceURI === DRAWING && ['r', 'br', 'fld'].includes(element.localName),
  );
  const textOf = (element: Element) =>
    element.localName === 'br' ? '\n' : (child(element, DRAWING, 't')?.textContent ?? '');
  const original = runs.map(textOf).join('');
  if (original === text) return;
  let start = 0;
  while (start < original.length && start < text.length && original[start] === text[start]) start++;
  let end = original.length;
  let replacementEnd = text.length;
  while (end > start && replacementEnd > start && original[end - 1] === text[replacementEnd - 1]) {
    end--;
    replacementEnd--;
  }

  const before: Element[] = [];
  const after: Element[] = [];
  let style: Element | undefined;
  let offset = 0;
  for (const run of runs) {
    const value = textOf(run);
    const next = offset + value.length;
    if (run.localName !== 'br' && offset <= start) style = child(run, DRAWING, 'rPr');
    const slice = (from: number, to: number) => {
      const clone = run.cloneNode(true) as Element;
      const content = child(clone, DRAWING, 't');
      if (content) content.textContent = value.slice(from, to);
      return clone;
    };
    if (offset < start) before.push(slice(0, Math.min(value.length, start - offset)));
    if (next > end) after.push(slice(Math.max(0, end - offset), value.length));
    offset = next;
  }

  const inserted: Element[] = [];
  text
    .slice(start, replacementEnd)
    .split('\n')
    .forEach((line, index) => {
      if (index) inserted.push(drawingElement(paragraph, 'br'));
      if (!line) return;
      const run = drawingElement(paragraph, 'r');
      if (style) run.append(style.cloneNode(true));
      const content = drawingElement(paragraph, 't');
      content.textContent = line;
      run.append(content);
      inserted.push(run);
    });
  const anchor = runs.at(-1)?.nextSibling ?? child(paragraph, DRAWING, 'endParaRPr') ?? null;
  for (const run of runs) run.remove();
  for (const run of [...before, ...inserted, ...after]) paragraph.insertBefore(run, anchor);
}

function replaceBounds(object: Element, node: SlideNode, bounds: Rect) {
  if (
    ![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) ||
    bounds.width <= 0 ||
    bounds.height <= 0
  )
    throw new Error('Enter a positive width and height.');
  const frame = object.localName === 'graphicFrame';
  const group = object.localName === 'grpSp';
  let properties = frame ? object : child(object, PRESENTATION, group ? 'grpSpPr' : 'spPr');
  if (!properties) throw new Error('This object does not support resizing.');
  const namespace = frame ? PRESENTATION : DRAWING;
  let transform = child(properties, namespace, 'xfrm');
  if (!transform) {
    transform = object.ownerDocument.createElementNS(namespace, frame ? 'p:xfrm' : 'a:xfrm');
    const anchor = frame ? (child(properties, DRAWING, 'graphic') ?? null) : properties.firstChild;
    properties.insertBefore(transform, anchor);
    if (node.transform.rotation)
      transform.setAttribute('rot', String(Math.round(node.transform.rotation * 60000)));
    if (node.transform.flipHorizontal) transform.setAttribute('flipH', '1');
    if (node.transform.flipVertical) transform.setAttribute('flipV', '1');
  }
  let off = child(transform, DRAWING, 'off');
  let ext = child(transform, DRAWING, 'ext');
  if (!off) {
    off = drawingElement(transform, 'off');
    transform.insertBefore(off, transform.firstChild);
  }
  if (!ext) {
    ext = drawingElement(transform, 'ext');
    off.after(ext);
  }
  off.setAttribute('x', String(Math.round(bounds.x)));
  off.setAttribute('y', String(Math.round(bounds.y)));
  ext.setAttribute('cx', String(Math.round(bounds.width)));
  ext.setAttribute('cy', String(Math.round(bounds.height)));
  if (group && node.type === 'group') {
    const coordinate = node.childTransform ?? { ...node.transform, x: 0, y: 0 };
    if (!child(transform, DRAWING, 'chOff')) {
      const value = drawingElement(transform, 'chOff');
      value.setAttribute('x', String(coordinate.x));
      value.setAttribute('y', String(coordinate.y));
      transform.append(value);
    }
    if (!child(transform, DRAWING, 'chExt')) {
      const value = drawingElement(transform, 'chExt');
      value.setAttribute('cx', String(coordinate.width));
      value.setAttribute('cy', String(coordinate.height));
      transform.append(value);
    }
  }
  if (node.type === 'table') {
    for (const column of object.getElementsByTagNameNS(DRAWING, 'gridCol')) {
      column.setAttribute(
        'w',
        String(
          Math.round((Number(column.getAttribute('w')) * bounds.width) / node.transform.width),
        ),
      );
    }
    for (const row of object.getElementsByTagNameNS(DRAWING, 'tr')) {
      row.setAttribute(
        'h',
        String(Math.round((Number(row.getAttribute('h')) * bounds.height) / node.transform.height)),
      );
    }
  }
}

export async function editPptx(bytes: Uint8Array, edit: ObjectEdit): Promise<Uint8Array> {
  const { node } = edit;
  if (!node.sourcePart || !node.sourceId) throw new Error('This object cannot be edited.');
  const parts = await new Promise<Record<string, Uint8Array>>((resolve, reject) => {
    unzip(bytes, (error, result) => (error ? reject(error) : resolve(result)));
  });
  const part = parts[node.sourcePart];
  if (!part) throw new Error('The original slide could not be found.');
  const xml = new DOMParser().parseFromString(strFromU8(part), 'application/xml');
  if (xml.getElementsByTagName('parsererror').length)
    throw new Error('The slide contains invalid XML.');
  const identifiers = Array.from(xml.getElementsByTagNameNS(PRESENTATION, 'cNvPr')).filter(
    (element) => element.getAttribute('id') === node.sourceId,
  );
  const object = identifiers.length === 1 ? identifiers[0]?.parentElement?.parentElement : null;
  if (!object) throw new Error('The original object could not be found.');
  const locks: ObjectLocks = { ...node.editing?.locks };
  const nonVisual = Array.from(object.children).find((element) =>
    element.localName.startsWith('nv'),
  );
  for (const element of nonVisual?.getElementsByTagNameNS(DRAWING, '*') ?? []) {
    if (!element.localName.endsWith('Locks')) continue;
    for (const attribute of element.attributes) {
      if (attribute.value === '1' || attribute.value === 'true')
        locks[attribute.localName as keyof ObjectLocks] = true;
    }
  }
  if (locks?.noSelection) throw new Error('This object is locked.');
  if (edit.paragraph) {
    if (locks?.noTextEdit) throw new Error('Text editing is locked for this object.');
    let body = child(object, PRESENTATION, 'txBody');
    if (
      node.type === 'table' &&
      edit.paragraph.row !== undefined &&
      edit.paragraph.column !== undefined
    ) {
      const table = object.getElementsByTagNameNS(DRAWING, 'tbl')[0];
      const row = table && children(table, DRAWING, 'tr')[edit.paragraph.row];
      const cell = row && children(row, DRAWING, 'tc')[edit.paragraph.column];
      body = cell && child(cell, DRAWING, 'txBody');
    }
    const paragraph = body && children(body, DRAWING, 'p')[edit.paragraph.index];
    if (!paragraph) throw new Error('This text cannot be edited.');
    replaceParagraphText(paragraph, edit.paragraph.text.replace(/\r\n?/g, '\n'));
  }
  if (edit.bounds) {
    const b = edit.bounds;
    if (locks?.noMove && (b.x !== node.transform.x || b.y !== node.transform.y))
      throw new Error('This object cannot be moved.');
    if (locks?.noResize && (b.width !== node.transform.width || b.height !== node.transform.height))
      throw new Error('This object cannot be resized.');
    if (
      locks.noChangeAspect &&
      Math.abs(b.width * node.transform.height - b.height * node.transform.width) >
        Math.max(node.transform.width, node.transform.height)
    )
      throw new Error('This object must keep its proportions.');
    replaceBounds(object, node, b);
  }
  parts[node.sourcePart] = strToU8(new XMLSerializer().serializeToString(xml));
  return new Promise((resolve, reject) => {
    zip(parts, { level: 6 }, (error, result) => (error ? reject(error) : resolve(result)));
  });
}
