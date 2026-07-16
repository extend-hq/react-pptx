import { describe, expect, it } from 'vitest';
import { PptxFontManager, resolvePptxFontFamily } from './fonts';

describe('PowerPoint font resolution', () => {
  it('adds Office-compatible and script-complete fallbacks', () => {
    const stack = resolvePptxFontFamily('Aptos');
    expect(stack).toContain('"Aptos"');
    expect(stack).toContain('"Segoe UI"');
    expect(stack).toContain('"Noto Sans CJK JP"');
    expect(stack).toContain('sans-serif');
    expect(stack.indexOf('"Segoe UI"')).toBeLessThan(stack.indexOf('sans-serif'));
  });

  it('does not let emoji fonts claim ordinary word spaces', () => {
    const stack = resolvePptxFontFamily('Roboto');

    expect(stack).not.toContain('Apple Color Emoji');
    expect(stack).not.toContain('Segoe UI Emoji');
    expect(stack).toContain('sans-serif');
  });

  it('supports host overrides without duplicating families', () => {
    const stack = resolvePptxFontFamily('Brand Sans, Arial', {
      fallbacks: { 'Brand Sans': ['Inter', 'Arial'] },
      fallbackFamilies: ['Inter'],
    });
    expect(stack.match(/"Inter"/g)).toHaveLength(1);
    expect(stack.match(/"Arial"/g)).toHaveLength(1);
  });

  it('keeps East Asian fonts ahead of generic fallbacks', () => {
    const stack = resolvePptxFontFamily('SimSun');
    expect(stack.indexOf('"Songti SC"')).toBeLessThan(stack.indexOf('"Noto Sans"'));
  });

  it('uses metric-appropriate fallbacks for the reported deck fonts', () => {
    const heading = resolvePptxFontFamily('Sitka Heading');
    const body = resolvePptxFontFamily('Source Sans Pro');

    expect(heading.indexOf('"Georgia"')).toBeLessThan(heading.indexOf('sans-serif'));
    expect(heading).toContain('serif');
    expect(body.indexOf('"Helvetica Neue"')).toBeLessThan(body.indexOf('"Noto Sans"'));
  });

  it('applies fallback stacks to HTML and SVG font declarations', () => {
    const root = document.createElement('div');
    const html = document.createElement('span');
    html.style.fontFamily = 'Calibri';
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    svg.setAttribute('font-family', 'MS Gothic');
    root.append(html, svg);

    new PptxFontManager({ reportMissingFonts: false }).applyTo(root);

    expect(html.style.fontFamily).toContain('Carlito');
    expect(svg.getAttribute('style')).toContain('Yu Gothic');
  });
});
