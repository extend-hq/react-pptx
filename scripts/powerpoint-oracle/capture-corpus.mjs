#!/usr/bin/env node
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { chromium } from '@playwright/test';
import { PNG } from 'pngjs';
import {
  assertRendererEnvironment,
  resolveCorpusPath,
  validateManifestValue,
} from './contract.mjs';
import { buildOfficeFontCatalog, matchOracleFonts } from './font-catalog.mjs';

const manifestPath = path.resolve(process.argv[2] ?? '.powerpoint-oracle/manifest.json');
const captureRoot = path.resolve(process.argv[3] ?? '.powerpoint-oracle/captures');
const baseUrl = process.argv[4] ?? 'http://127.0.0.1:4173';
const requestedLimit = Number(process.argv[5] ?? 0);
const limit = Number.isFinite(requestedLimit) && requestedLimit > 0 ? requestedLimit : Infinity;
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
const contractFailures = validateManifestValue(manifest);
if (contractFailures.length) throw new Error(contractFailures.join('\n'));

function cssString(value) {
  return JSON.stringify(value);
}

function mimeType(file) {
  const extension = path.extname(file).toLowerCase();
  if (extension === '.otf') return 'font/otf';
  if (extension === '.ttc') return 'font/collection';
  return 'font/ttf';
}

async function fontCss(entries) {
  const rules = [];
  for (const entry of entries) {
    const bytes = await readFile(entry.file);
    rules.push(
      `@font-face{font-family:${cssString(entry.family)};src:url(data:${mimeType(entry.file)};base64,${bytes.toString('base64')});font-style:${entry.style};font-weight:${entry.weight};font-display:block}`,
    );
  }
  return rules.join('\n');
}

const { entries: fontCatalog, fingerprint } = await buildOfficeFontCatalog();
const browser = await chromium.launch({ headless: true });
const context = await browser.newContext({
  viewport: manifest.expectedRendererEnvironment.viewport,
  deviceScaleFactor: manifest.expectedRendererEnvironment.deviceScaleFactor,
  locale: manifest.expectedRendererEnvironment.locale ?? 'en-US',
  timezoneId: manifest.expectedRendererEnvironment.timezone ?? 'UTC',
});
assertRendererEnvironment(manifest.expectedRendererEnvironment, {
  browserName: 'chromium',
  browserVersion: browser.version(),
  platform: process.platform,
  architecture: process.arch,
  viewport: manifest.expectedRendererEnvironment.viewport,
  deviceScaleFactor: manifest.expectedRendererEnvironment.deviceScaleFactor,
  locale: manifest.expectedRendererEnvironment.locale ?? 'en-US',
  timezone: manifest.expectedRendererEnvironment.timezone ?? 'UTC',
  fontFingerprint: fingerprint,
});
const page = await context.newPage();
await page.goto(baseUrl);
const failures = [];

for (const [deckIndex, deck] of manifest.decks.slice(0, limit).entries()) {
  process.stdout.write(`[${deckIndex + 1}/${Math.min(manifest.decks.length, limit)}] ${deck.id}\n`);
  try {
    await page.reload();
    const matchedFonts = matchOracleFonts(deck.oracleFonts ?? [], fontCatalog);
    if (matchedFonts.length) {
      await page.addStyleTag({ content: await fontCss(matchedFonts) });
      await page.evaluate(
        async (families) => {
          await Promise.all(
            families.map((family) => document.fonts.load(`16px ${JSON.stringify(family)}`)),
          );
          await document.fonts.ready;
        },
        [...new Set(matchedFonts.map((entry) => entry.family))],
      );
    }

    const source = resolveCorpusPath(manifestPath, deck.source);
    await page.locator('input[type="file"]').setInputFiles(source);
    await page.getByText(path.basename(source), { exact: true }).waitFor({ timeout: 60_000 });
    await page.getByRole('button', { name: 'Single slide' }).click();
    const outputDirectory = path.join(captureRoot, deck.id);
    await mkdir(outputDirectory, { recursive: true });

    for (const slide of deck.slides) {
      const slideNumber = slide.index + 1;
      const reference = PNG.sync.read(
        await readFile(resolveCorpusPath(manifestPath, slide.reference)),
      );
      await page.getByText(`${slideNumber} / ${deck.slides.length}`, { exact: true }).waitFor();
      const slideElement = page.locator('.rpv-stage > .rpv-viewport > div > div');
      await slideElement.waitFor({ state: 'visible', timeout: 30_000 });
      const clip = await slideElement.evaluate(
        async (element, expectedSize) => {
          await document.fonts.ready;
          await Promise.all(
            [...element.querySelectorAll('img')].map((image) =>
              image.complete
                ? undefined
                : new Promise((resolve) => {
                    image.addEventListener('load', resolve, { once: true });
                    image.addEventListener('error', resolve, { once: true });
                  }),
            ),
          );
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          const frame = element.parentElement;
          const indexWrapper = frame?.parentElement;
          element.style.transform = 'none';
          element.style.transformOrigin = 'left top';
          if (frame) {
            frame.style.width = `${expectedSize.width}px`;
            frame.style.height = `${expectedSize.height}px`;
            frame.style.boxShadow = 'none';
          }
          if (indexWrapper) indexWrapper.style.width = 'fit-content';
          const bounds = element.getBoundingClientRect();
          return {
            x: bounds.x,
            y: bounds.y,
            width: expectedSize.width,
            height: expectedSize.height,
          };
        },
        { width: reference.width, height: reference.height },
      );
      await page.screenshot({
        path: path.join(outputDirectory, `slide-${slideNumber}.png`),
        animations: 'disabled',
        scale: 'css',
        clip,
      });
      if (slideNumber < deck.slides.length) {
        await page.getByRole('button', { name: 'Next slide' }).click();
      }
    }
  } catch (error) {
    failures.push({
      deck: deck.id,
      source: deck.source,
      error: error instanceof Error ? error.message : String(error),
    });
    process.stderr.write(`  failed: ${failures.at(-1).error}\n`);
  }
}

await context.close();
await browser.close();
await mkdir(captureRoot, { recursive: true });
await writeFile(
  path.join(captureRoot, 'capture-failures.json'),
  `${JSON.stringify(failures, null, 2)}\n`,
);
console.log(
  `Captured ${manifest.decks.slice(0, limit).length - failures.length}/${manifest.decks.slice(0, limit).length} deck(s) into ${captureRoot}.`,
);
if (failures.length) process.exitCode = 1;
