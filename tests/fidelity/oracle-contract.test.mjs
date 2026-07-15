import assert from 'node:assert/strict';
import test from 'node:test';
import {
  assertRendererEnvironment,
  validateManifestValue,
} from '../../scripts/powerpoint-oracle/contract.mjs';

const environment = {
  browserName: 'chromium',
  browserVersion: '1',
  viewport: { width: 1440, height: 900 },
  deviceScaleFactor: 1,
  fontFingerprint: 'abc',
};

test('oracle manifest requires stable renderer environment and slide thresholds', () => {
  assert.deepEqual(
    validateManifestValue({
      version: 1,
      expectedRendererEnvironment: environment,
      decks: [
        {
          id: 'basic',
          source: 'basic.pptx',
          maxPixelDifferenceRatio: 0.02,
          slides: [{ index: 0, reference: 'basic.png' }],
        },
      ],
    }),
    [],
  );
});

test('oracle rejects missing environment pins', () => {
  const failures = validateManifestValue({ version: 1, decks: [] });
  assert.ok(failures.some((failure) => failure.includes('fontFingerprint')));
});

test('renderer environment comparison fails closed', () => {
  assert.throws(
    () => assertRendererEnvironment(environment, { ...environment, browserVersion: '2' }),
    /browserVersion/,
  );
});
