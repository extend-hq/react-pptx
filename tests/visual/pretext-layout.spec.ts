import { expect, test } from '@playwright/test';

test('renders the Pretext text-layout regression fixture without viewer errors', async ({
  page,
}) => {
  const consoleErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });

  await page.goto('/');
  await page.locator('input[type=file]').setInputFiles('tests/fixtures/pretext-layout-cases.pptx');
  await expect(page.getByRole('contentinfo')).toContainText('Slide 1 of 8');
  await page.getByRole('contentinfo').getByRole('button', { name: 'Normal', exact: true }).click();

  const markers = [
    'Pretext layout',
    'Justification exposes the word-spacing gap',
    'Formatting runs should not move the words',
    'Tabs and hanging indents require paragraph geometry',
    'Script-aware wrapping changes the break opportunities',
    'Columns and autofit depend on measured line results',
    'A synthetic image exclusion shows the Pretext opportunity',
    'The words finish',
  ];
  const viewport = page.getByTestId('pptx-viewport');
  for (const [index, marker] of markers.entries()) {
    await expect(viewport.getByRole('region', { name: `Slide ${index + 1} of 8` })).toContainText(
      marker,
    );
    if (index < markers.length - 1) await page.getByRole('button', { name: 'Next slide' }).click();
  }
  expect(consoleErrors).toEqual([]);
});
