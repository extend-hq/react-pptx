import type { SlideNode } from '@extend-ai/react-pptx-model';
import type { ObjectEdit } from './pptx-editing';

export type InlineTextSession = { body: HTMLElement; finish: (save?: boolean) => void };

function readText(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent ?? '';
  if (!(node instanceof HTMLElement)) return '';
  if (node.hasAttribute('data-rpv-bullet')) return '';
  if (node.tagName === 'BR') return '\n';
  return Array.from(node.childNodes).map(readText).join('');
}

function paragraphsFrom(body: HTMLElement): NonNullable<ObjectEdit['text']>['paragraphs'] {
  const paragraphs: NonNullable<ObjectEdit['text']>['paragraphs'] = [];
  let loose = '';
  for (const child of body.childNodes) {
    if (child instanceof HTMLElement && ['DIV', 'P'].includes(child.tagName)) {
      if (loose) {
        paragraphs.push({ text: loose });
        loose = '';
      }
      const source = child.dataset.rpvTextParagraph;
      let text = readText(child);
      if (child.lastChild instanceof HTMLBRElement) text = text.slice(0, -1);
      paragraphs.push({ text, ...(source === undefined ? {} : { sourceIndex: Number(source) }) });
    } else loose += readText(child);
  }
  if (loose || !paragraphs.length) paragraphs.push({ text: loose === '\n' ? '' : loose });
  return paragraphs;
}

export function startInlineTextEdit(
  body: HTMLElement,
  node: SlideNode,
  onSave: (edit: ObjectEdit) => Promise<void>,
  onFinish: () => void,
): InlineTextSession {
  const originalHtml = body.innerHTML;
  const originalStyle = body.getAttribute('style');
  const initial = paragraphsFrom(body).map(({ text }) => text);
  const cell = body.closest<HTMLElement>('[data-rpv-cell-row]');
  let finished = false;
  body.contentEditable = 'true';
  body.setAttribute('role', 'textbox');
  body.setAttribute('aria-label', 'Slide text');
  body.setAttribute('aria-multiline', 'true');
  body.style.outline = 'none';
  body.style.cursor = 'text';
  body.style.userSelect = 'text';
  body.style.minHeight = '1em';
  for (const marker of body.querySelectorAll<HTMLElement>('[data-rpv-bullet]')) {
    marker.contentEditable = 'false';
    marker.style.userSelect = 'none';
  }
  const finish = (save = true) => {
    if (finished) return;
    finished = true;
    const paragraphs = paragraphsFrom(body);
    body.removeEventListener('blur', blur);
    body.removeEventListener('keydown', keydown);
    body.removeEventListener('paste', paste);
    body.removeEventListener('beforeinput', beforeInput);
    body.removeAttribute('contenteditable');
    body.removeAttribute('role');
    body.removeAttribute('aria-label');
    body.removeAttribute('aria-multiline');
    if (originalStyle === null) body.removeAttribute('style');
    else body.setAttribute('style', originalStyle);
    if (!save) body.innerHTML = originalHtml;
    onFinish();
    if (save && JSON.stringify(initial) !== JSON.stringify(paragraphs.map(({ text }) => text))) {
      void onSave({
        node,
        text: {
          paragraphs,
          ...(cell
            ? { row: Number(cell.dataset.rpvCellRow), column: Number(cell.dataset.rpvCellColumn) }
            : {}),
        },
      });
    }
  };
  const blur = () => finish();
  const keydown = (event: KeyboardEvent) => {
    event.stopPropagation();
    if (event.key === 'Escape' || (event.key === 'Enter' && (event.ctrlKey || event.metaKey))) {
      event.preventDefault();
      finish();
      body.closest<HTMLElement>('[data-rpv-node-path]')?.focus({ preventScroll: true });
    }
  };
  const paste = (event: ClipboardEvent) => {
    event.preventDefault();
    document.execCommand('insertText', false, event.clipboardData?.getData('text/plain') ?? '');
  };
  const beforeInput = (event: InputEvent) => {
    if (event.inputType.startsWith('format')) event.preventDefault();
  };
  body.addEventListener('blur', blur);
  body.addEventListener('keydown', keydown);
  body.addEventListener('paste', paste);
  body.addEventListener('beforeinput', beforeInput);
  return { body, finish };
}
