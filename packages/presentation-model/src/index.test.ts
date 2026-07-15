import { describe, expect, it } from 'vitest';
import { collectSlideText, type PresentationSlide } from './index';

describe('collectSlideText', () => {
  it('collects nested shape, group, table, and chart text', () => {
    const slide: PresentationSlide = {
      id: 's1',
      index: 0,
      nodes: [
        {
          id: 'shape',
          type: 'shape',
          transform: { x: 0, y: 0, width: 10, height: 10 },
          geometry: { preset: 'rect' },
          paragraphs: [{ runs: [{ text: 'Quarterly' }] }],
        },
        {
          id: 'group',
          type: 'group',
          transform: { x: 0, y: 0, width: 10, height: 10 },
          children: [
            {
              id: 'chart',
              type: 'chart',
              transform: { x: 0, y: 0, width: 10, height: 10 },
              chartType: 'bar',
              title: 'Revenue',
              series: [{ name: 'North America', values: [1] }],
            },
          ],
        },
      ],
    };

    expect(collectSlideText(slide)).toBe('Quarterly\nRevenue\nNorth America');
  });
});
