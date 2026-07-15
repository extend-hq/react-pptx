import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import path from 'node:path';

export const MANIFEST_VERSION = 1;

export function validateManifestValue(manifest) {
  const failures = [];
  if (!manifest || typeof manifest !== 'object') failures.push('manifest must be an object');
  if (manifest?.version !== MANIFEST_VERSION) failures.push(`version must be ${MANIFEST_VERSION}`);
  if (!Array.isArray(manifest?.decks) || manifest.decks.length === 0) {
    failures.push('decks must be a non-empty array');
  }
  const ids = new Set();
  for (const [index, deck] of (manifest?.decks ?? []).entries()) {
    const at = `decks[${index}]`;
    if (!deck.id || typeof deck.id !== 'string') failures.push(`${at}.id must be a string`);
    if (ids.has(deck.id)) failures.push(`${at}.id duplicates ${deck.id}`);
    ids.add(deck.id);
    if (!deck.source || typeof deck.source !== 'string')
      failures.push(`${at}.source must be a path`);
    if (!Array.isArray(deck.slides) || deck.slides.length === 0)
      failures.push(`${at}.slides must not be empty`);
    for (const [slideIndex, slide] of (deck.slides ?? []).entries()) {
      const slideAt = `${at}.slides[${slideIndex}]`;
      if (!Number.isInteger(slide.index) || slide.index < 0)
        failures.push(`${slideAt}.index must be >= 0`);
      if (!slide.reference) failures.push(`${slideAt}.reference is required`);
      const threshold = slide.maxPixelDifferenceRatio ?? deck.maxPixelDifferenceRatio;
      if (typeof threshold !== 'number' || threshold < 0 || threshold > 1) {
        failures.push(`${slideAt} needs maxPixelDifferenceRatio in [0, 1]`);
      }
    }
  }
  const environment = manifest?.expectedRendererEnvironment;
  for (const key of [
    'browserName',
    'browserVersion',
    'viewport',
    'deviceScaleFactor',
    'fontFingerprint',
  ]) {
    if (environment?.[key] === undefined)
      failures.push(`expectedRendererEnvironment.${key} is required`);
  }
  return failures;
}

export async function sha256File(filePath) {
  const bytes = await readFile(filePath);
  return createHash('sha256').update(bytes).digest('hex');
}

export function resolveCorpusPath(manifestPath, value) {
  return path.resolve(path.dirname(manifestPath), value);
}

export function assertRendererEnvironment(expected, actual) {
  const mismatches = [];
  for (const key of [
    'browserName',
    'browserVersion',
    'platform',
    'architecture',
    'locale',
    'timezone',
    'deviceScaleFactor',
    'fontFingerprint',
  ]) {
    if (expected[key] !== undefined && expected[key] !== actual[key]) {
      mismatches.push(
        `${key}: expected ${JSON.stringify(expected[key])}, got ${JSON.stringify(actual[key])}`,
      );
    }
  }
  if (expected.viewport) {
    for (const key of ['width', 'height']) {
      if (expected.viewport[key] !== actual.viewport?.[key]) {
        mismatches.push(
          `viewport.${key}: expected ${expected.viewport[key]}, got ${actual.viewport?.[key]}`,
        );
      }
    }
  }
  if (mismatches.length)
    throw new Error(`Renderer environment mismatch:\n${mismatches.join('\n')}`);
}
