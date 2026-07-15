#!/usr/bin/env node
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { resolveCorpusPath, validateManifestValue } from './contract.mjs';

const manifestPath = path.resolve(process.argv[2] ?? '.powerpoint-oracle/manifest.json');
const captureRoot = path.resolve(process.argv[3] ?? '.powerpoint-oracle/captures');
const outputRoot = path.resolve(process.argv[4] ?? '.powerpoint-oracle/results');
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
const contractFailures = validateManifestValue(manifest);
if (contractFailures.length) throw new Error(contractFailures.join('\n'));

await mkdir(outputRoot, { recursive: true });
const results = [];
for (const deck of manifest.decks) {
  for (const slide of deck.slides) {
    const referencePath = resolveCorpusPath(manifestPath, slide.reference);
    const actualPath = path.join(captureRoot, deck.id, `slide-${slide.index + 1}.png`);
    const reference = PNG.sync.read(await readFile(referencePath));
    let actual;
    try {
      actual = PNG.sync.read(await readFile(actualPath));
    } catch (error) {
      results.push({
        deck: deck.id,
        slide: slide.index,
        pass: false,
        reason: 'missing-capture',
        error: error instanceof Error ? error.message : String(error),
      });
      continue;
    }
    if (reference.width !== actual.width || reference.height !== actual.height) {
      results.push({
        deck: deck.id,
        slide: slide.index,
        pass: false,
        reason: 'dimensions',
        reference: [reference.width, reference.height],
        actual: [actual.width, actual.height],
      });
      continue;
    }
    const diff = new PNG({ width: reference.width, height: reference.height });
    const changed = pixelmatch(
      reference.data,
      actual.data,
      diff.data,
      reference.width,
      reference.height,
      {
        threshold: slide.pixelThreshold ?? deck.pixelThreshold ?? 0.1,
        includeAA: false,
      },
    );
    const ratio = changed / (reference.width * reference.height);
    const maximum = slide.maxPixelDifferenceRatio ?? deck.maxPixelDifferenceRatio;
    const pass = ratio <= maximum;
    const result = {
      deck: deck.id,
      slide: slide.index,
      pass,
      changedPixels: changed,
      ratio,
      maximum,
    };
    results.push(result);
    if (!pass) {
      const dir = path.join(outputRoot, deck.id);
      await mkdir(dir, { recursive: true });
      await writeFile(path.join(dir, `slide-${slide.index + 1}-diff.png`), PNG.sync.write(diff));
    }
  }
}

await writeFile(path.join(outputRoot, 'summary.json'), `${JSON.stringify(results, null, 2)}\n`);
const failed = results.filter((result) => !result.pass);
const compared = results.filter((result) => typeof result.ratio === 'number');
const ratios = compared.map((result) => result.ratio).sort((a, b) => a - b);
const percentile = (value) =>
  ratios[Math.min(ratios.length - 1, Math.floor(ratios.length * value))];
const metrics = {
  manifestDecks: manifest.decks.length,
  manifestSlides: results.length,
  comparedSlides: compared.length,
  passedSlides: compared.filter((result) => result.pass).length,
  failedSlides: compared.filter((result) => !result.pass).length,
  excludedSlides: results.length - compared.length,
  meanRatio: ratios.length ? ratios.reduce((sum, ratio) => sum + ratio, 0) / ratios.length : null,
  medianRatio: ratios.length ? percentile(0.5) : null,
  p90Ratio: ratios.length ? percentile(0.9) : null,
  p95Ratio: ratios.length ? percentile(0.95) : null,
  maximumRatio: ratios.at(-1) ?? null,
};
await writeFile(path.join(outputRoot, 'metrics.json'), `${JSON.stringify(metrics, null, 2)}\n`);
console.log(
  `Compared ${metrics.comparedSlides}/${metrics.manifestSlides} slide(s); ${metrics.failedSlides} exceeded the visual threshold and ${metrics.excludedSlides} lacked aligned captures.`,
);
if (failed.length) process.exitCode = 1;
