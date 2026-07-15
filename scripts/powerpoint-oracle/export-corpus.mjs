#!/usr/bin/env node
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdir, readdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { chromium } from '@playwright/test';
import { buildOfficeFontCatalog } from './font-catalog.mjs';
import { sha256File } from './contract.mjs';

const corpusArgument = process.argv[2] ?? process.env.POWERPOINT_ORACLE_CORPUS;
if (!corpusArgument) {
  throw new Error(
    'Usage: pnpm powerpoint-oracle:export <corpus-path> [output-path] [limit], or set POWERPOINT_ORACLE_CORPUS',
  );
}
const corpusRoot = path.resolve(corpusArgument);
const oracleRoot = path.resolve(process.argv[3] ?? '.powerpoint-oracle');
const requestedLimit = Number(process.argv[4] ?? 0);
const limit = Number.isFinite(requestedLimit) && requestedLimit > 0 ? requestedLimit : Infinity;
const referencesRoot = path.join(oracleRoot, 'references');
const pdfRoot = path.join(oracleRoot, 'pdf');
const scriptPath = path.resolve('scripts/powerpoint-oracle/export_pptx_to_pdf.applescript');

async function listDecks(directory) {
  if ((await stat(directory)).isFile()) return [directory];
  const decks = [];
  const visit = async (current) => {
    for (const entry of await readdir(current, { withFileTypes: true })) {
      const absolute = path.join(current, entry.name);
      if (entry.isDirectory()) {
        if (entry.name !== 'invalid-parser-fixtures') await visit(absolute);
      } else if (entry.isFile() && entry.name.toLowerCase().endsWith('.pptx')) decks.push(absolute);
    }
  };
  await visit(directory);
  return decks.sort().slice(0, limit);
}

function deckId(source) {
  const relative = (
    corpusRoot === source ? path.basename(source) : path.relative(corpusRoot, source)
  ).replace(/\.pptx$/i, '');
  const slug = relative
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/gi, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 96)
    .toLowerCase();
  const suffix = createHash('sha1').update(relative).digest('hex').slice(0, 8);
  return `${slug}-${suffix}`;
}

function pageCount(pdfPath) {
  const output = execFileSync('pdfinfo', [pdfPath], { encoding: 'utf8' });
  const match = output.match(/^Pages:\s+(\d+)$/m);
  if (!match) throw new Error(`Unable to read page count from ${pdfPath}`);
  return Number(match[1]);
}

function oracleFonts(pdfPath) {
  const output = execFileSync('pdffonts', [pdfPath], { encoding: 'utf8' });
  return [
    ...new Set(
      output
        .split('\n')
        .slice(2)
        .map((line) => line.slice(0, 37).trim())
        .filter(Boolean),
    ),
  ].sort();
}

async function renderPdfPages(pdfPath, outputDirectory, pages) {
  await rm(outputDirectory, { recursive: true, force: true });
  await mkdir(outputDirectory, { recursive: true });
  const prefix = path.join(outputDirectory, 'raw');
  execFileSync('pdftoppm', ['-png', '-r', '96', pdfPath, prefix], { stdio: 'ignore' });
  const generated = (await readdir(outputDirectory))
    .filter((name) => /^raw-\d+\.png$/.test(name))
    .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
  if (generated.length !== pages) {
    throw new Error(`Expected ${pages} PNGs from ${pdfPath}, got ${generated.length}`);
  }
  for (const [index, name] of generated.entries()) {
    await rename(
      path.join(outputDirectory, name),
      path.join(outputDirectory, `slide-${index + 1}.png`),
    );
  }
}

async function exportPdf(source, output) {
  await mkdir(path.dirname(output), { recursive: true });
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      await rm(output, { force: true });
      execFileSync('osascript', [scriptPath, source, output], {
        encoding: 'utf8',
        timeout: 120_000,
        stdio: ['ignore', 'pipe', 'pipe'],
      });
      const bytes = await readFile(output);
      if (bytes.length === 0) throw new Error('PowerPoint created an empty PDF');
      return;
    } catch (error) {
      if (attempt === 3) throw error;
    }
  }
}

await mkdir(referencesRoot, { recursive: true });
await mkdir(pdfRoot, { recursive: true });
const sources = await listDecks(corpusRoot);
const browser = await chromium.launch({ headless: true });
const browserVersion = browser.version();
await browser.close();

const decks = [];
const failures = [];
for (const [deckIndex, source] of sources.entries()) {
  const id = deckId(source);
  const pdfPath = path.join(pdfRoot, `${id}.pdf`);
  const referenceDirectory = path.join(referencesRoot, id);
  process.stdout.write(
    `[${deckIndex + 1}/${sources.length}] ${path.relative(corpusRoot, source)}\n`,
  );
  try {
    await exportPdf(source, pdfPath);
    const pages = pageCount(pdfPath);
    await renderPdfPages(pdfPath, referenceDirectory, pages);
    decks.push({
      id,
      source,
      sha256: await sha256File(source),
      oracleFonts: oracleFonts(pdfPath),
      maxPixelDifferenceRatio: 0.05,
      slides: Array.from({ length: pages }, (_, index) => ({
        index,
        reference: path.relative(
          oracleRoot,
          path.join(referenceDirectory, `slide-${index + 1}.png`),
        ),
      })),
    });
  } catch (error) {
    failures.push({ source, error: error instanceof Error ? error.message : String(error) });
    process.stderr.write(`  failed: ${failures.at(-1).error}\n`);
  }
}

const manifest = {
  version: 1,
  expectedRendererEnvironment: {
    browserName: 'chromium',
    browserVersion,
    platform: process.platform,
    architecture: process.arch,
    viewport: { width: 1800, height: 1400 },
    deviceScaleFactor: 1,
    locale: 'en-US',
    timezone: 'UTC',
    fontFingerprint: (await buildOfficeFontCatalog()).fingerprint,
  },
  decks,
};
await writeFile(path.join(oracleRoot, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
await writeFile(
  path.join(oracleRoot, 'export-failures.json'),
  `${JSON.stringify(failures, null, 2)}\n`,
);
console.log(
  `Exported ${decks.length}/${sources.length} decks (${decks.reduce((n, deck) => n + deck.slides.length, 0)} slides).`,
);
if (failures.length) process.exitCode = 1;
