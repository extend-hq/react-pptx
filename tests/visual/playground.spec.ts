import { expect, test } from '@playwright/test';

test('renders the generated PPTX and exposes viewer controls', async ({ page }) => {
  const consoleErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  await page.goto('/');
  await expect(page.getByText('PPTX · 2 slides')).toBeVisible();
  const titleFontStack = await page
    .locator('span')
    .filter({ hasText: /^A deterministic PowerPoint fixture$/ })
    .last()
    .evaluate((element) => element.style.fontFamily);
  expect(titleFontStack).toContain('Aptos Display');
  expect(titleFontStack).toContain('Noto Sans CJK JP');
  await page.getByRole('button', { name: 'Slides' }).click();
  await expect(page.getByRole('button', { name: 'Go to slide 1' })).toBeVisible();
  await expect(page.getByText('A deterministic PowerPoint fixture').last()).toBeVisible();

  await page.getByRole('textbox', { name: 'Search slide text' }).fill('Measured');
  await expect(page.getByText('1 hits')).toBeVisible();
  await page.getByRole('button', { name: 'Next search result' }).click();
  await expect(page.getByRole('group', { name: 'Slide navigation' })).toContainText('2 / 2');

  await page.getByRole('button', { name: 'Single slide' }).click();
  await expect(page.getByText('Cached navigation').last()).toBeVisible();
  expect(consoleErrors).toEqual([]);
});

test('shows owned parser diagnostics and accepts file uploads', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Diagnostics' }).click();
  const diagnostics = page.getByRole('complementary', { name: 'Rendering diagnostics' });
  await expect(diagnostics).toContainText('PPTX');
  await expect(diagnostics).toContainText('Slides2');

  await page
    .locator('input[type=file]')
    .setInputFiles('tests/fixtures/generated/viewer-smoke.pptx');
  await expect(page.getByText('PPTX · 2 slides')).toBeVisible();
});

test('loads and renders a real-world legacy PPT file', async ({ page }) => {
  const consoleErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });

  await page.goto('/');
  await page
    .locator('input[type=file]')
    .setInputFiles('tests/fixtures/legacy/file-example-250kb.ppt');

  await expect(page.getByText('PPT · 3 slides')).toBeVisible();
  await expect(page.getByRole('region', { name: 'Slide 1 of 3' })).toContainText('Lorem ipsum');
  await expect(page.getByTestId('pptx-viewport')).not.toContainText(
    'Click to edit the title text format',
  );

  await page.getByRole('button', { name: 'Single slide' }).click();
  await page.getByRole('button', { name: 'Next slide' }).click();
  await expect(page.getByRole('group', { name: 'Slide navigation' })).toContainText('2 / 3');
  const chartPreview = page
    .getByRole('region', { name: 'Slide 2 of 3' })
    .getByRole('img', { name: 'Legacy picture 5122' });
  await expect(chartPreview).toHaveAttribute('src', /^data:image\/png;base64,/);
  await expect
    .poll(() =>
      chartPreview.evaluate(
        (image: HTMLImageElement) =>
          image.complete && image.naturalWidth > 0 && image.naturalHeight > 0,
      ),
    )
    .toBe(true);
  const chartDimensions = await chartPreview.evaluate((image: HTMLImageElement) => ({
    width: image.naturalWidth,
    height: image.naturalHeight,
  }));
  expect(chartDimensions.width).toBeGreaterThan(0);
  expect(chartDimensions.height).toBeGreaterThan(0);

  await page.getByRole('button', { name: 'Next slide' }).click();
  await expect(page.getByRole('group', { name: 'Slide navigation' })).toContainText('3 / 3');
  await expect(page.getByRole('region', { name: 'Slide 3 of 3' })).toContainText('Table');

  await page.getByRole('button', { name: 'Slides' }).click();
  await expect(page.getByRole('button', { name: 'Go to slide 3' })).toBeVisible();

  await page.getByRole('button', { name: 'Diagnostics' }).click();
  const diagnostics = page.getByRole('complementary', { name: 'Rendering diagnostics' });
  await expect(diagnostics).toContainText('PPT');
  await expect(diagnostics).toContainText('degraded-rendering');
  expect(consoleErrors).toEqual([]);
});

test('scrolls the virtualized list smoothly without snapping back', async ({ page }) => {
  const consoleErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  await page.goto('/');
  await expect(page.getByText('PPTX · 2 slides')).toBeVisible();
  const viewport = page.getByTestId('pptx-viewport');
  await expect(page.getByRole('group', { name: 'Slide navigation' })).toContainText('1 / 2');
  // Fixed absolute offsets inside a full-height sizer keep the scroll
  // geometry stable while slides mount and unmount.
  await expect(viewport.locator('[data-rpv-virtual-sizer]')).toHaveCount(1);

  // Scroll to the bottom; the toolbar must follow the reading position.
  await viewport.evaluate((element) => {
    element.scrollTop = element.scrollHeight;
  });
  await expect(page.getByRole('group', { name: 'Slide navigation' })).toContainText('2 / 2');

  // The viewport must hold an arbitrary mid-scroll position: controlled
  // hosts echo onSlideChange back into slideIndex, and that echo used to
  // snap the scroll position to the slide boundary.
  await viewport.evaluate((element) => {
    element.scrollTop = 130;
  });
  await page.waitForTimeout(700);
  const settled = await viewport.evaluate((element) => element.scrollTop);
  expect(Math.round(settled)).toBe(130);
  expect(consoleErrors).toEqual([]);
});
