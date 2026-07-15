#!/usr/bin/env node
import { access, readFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { resolveCorpusPath, sha256File, validateManifestValue } from './contract.mjs';

const manifestPath = path.resolve(process.argv[2] ?? '.powerpoint-oracle/manifest.json');
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
const failures = validateManifestValue(manifest);

for (const deck of manifest.decks ?? []) {
  for (const [kind, relative] of [
    ['source', deck.source],
    ...deck.slides.map((slide) => [`slide ${slide.index} reference`, slide.reference]),
  ]) {
    try {
      await access(resolveCorpusPath(manifestPath, relative));
    } catch {
      failures.push(`${deck.id} ${kind} is missing: ${relative}`);
    }
  }
  if (deck.sha256) {
    const actual = await sha256File(resolveCorpusPath(manifestPath, deck.source));
    if (actual !== deck.sha256) failures.push(`${deck.id} source sha256 mismatch`);
  }
}

if (failures.length) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Validated ${manifest.decks.length} deck(s) in ${manifestPath}`);
}
