#!/usr/bin/env node

import { execFileSync } from 'node:child_process';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

const workspaceRoot = resolve(import.meta.dirname, '..');
const packageDirectory = join(workspaceRoot, 'packages/react-viewer');
const temporaryDirectory = mkdtempSync(join(tmpdir(), 'react-pptx-pack-'));

try {
  const output = execFileSync('npm', ['pack', '--json', '--pack-destination', temporaryDirectory], {
    cwd: packageDirectory,
    encoding: 'utf8',
  });
  const [packResult] = JSON.parse(output);
  if (!packResult?.filename) {
    throw new Error('npm pack did not report a tarball filename');
  }

  const tarball = join(temporaryDirectory, packResult.filename);
  execFileSync('tar', ['-xzf', tarball, '-C', temporaryDirectory]);

  const packedDirectory = join(temporaryDirectory, 'package');
  const manifest = JSON.parse(readFileSync(join(packedDirectory, 'package.json'), 'utf8'));
  const dependencyGroups = [
    'dependencies',
    'optionalDependencies',
    'peerDependencies',
    'devDependencies',
  ];

  for (const group of dependencyGroups) {
    for (const [name, version] of Object.entries(manifest[group] ?? {})) {
      if (String(version).startsWith('workspace:')) {
        throw new Error(`Packed manifest leaks workspace dependency ${group}.${name}`);
      }
      if (name.startsWith('@extend-ai/react-pptx-')) {
        throw new Error(`Packed manifest exposes internal package ${group}.${name}`);
      }
    }
  }

  // The chart stack ships as ordinary runtime dependencies; only this exact
  // set is allowed so stray packages cannot creep into the manifest.
  const allowedRuntimeDependencies = new Set([
    '@tanstack/virtual-core',
    'd3-geo',
    'd3-hierarchy',
    'd3-scale',
    'd3-shape',
    'regl',
    'topojson-client',
  ]);
  const runtimeDependencies = new Set([
    ...Object.keys(manifest.dependencies ?? {}),
    ...Object.keys(manifest.optionalDependencies ?? {}),
  ]);
  const unexpectedDependencies = [...runtimeDependencies].filter(
    (name) => !allowedRuntimeDependencies.has(name),
  );
  if (unexpectedDependencies.length > 0) {
    throw new Error(
      `Packed package unexpectedly requires runtime packages: ${unexpectedDependencies.join(', ')}`,
    );
  }
  const missingDependencies = [...allowedRuntimeDependencies].filter(
    (name) => !runtimeDependencies.has(name),
  );
  if (missingDependencies.length > 0) {
    throw new Error(`Packed package is missing runtime dependencies: ${missingDependencies.join(', ')}`);
  }
  for (const atlasPackage of ['us-atlas', 'world-atlas']) {
    if (runtimeDependencies.has(atlasPackage)) {
      throw new Error(`${atlasPackage} must stay bundled in the lazy atlas chunk, not a dependency`);
    }
  }

  const expectedPeers = new Set(['react', 'react-dom']);
  const actualPeers = Object.keys(manifest.peerDependencies ?? {});
  if (
    actualPeers.length !== expectedPeers.size ||
    actualPeers.some((dependency) => !expectedPeers.has(dependency))
  ) {
    throw new Error(`Packed package has unexpected peer dependencies: ${actualPeers.join(', ')}`);
  }

  const requiredFiles = [
    'dist/index.js',
    'dist/index.d.ts',
    'dist/index.css',
    'dist/native-parser-worker.js',
    'dist/pptx_wasm_bg.wasm',
    'README.md',
    'CHANGELOG.md',
    'LICENSE',
    'THIRD_PARTY_NOTICES.md',
    'THIRD_PARTY_LICENSES/Apache-2.0.txt',
  ];

  for (const file of requiredFiles) {
    if (!existsSync(join(packedDirectory, file))) {
      throw new Error(`Packed package is missing ${file}`);
    }
  }

  // The multi-megabyte TopoJSON atlases must live in a code-split chunk that
  // only loads when a presentation contains an Excel map chart.
  const distFiles = readdirSync(join(packedDirectory, 'dist')).filter((file) =>
    file.endsWith('.js'),
  );
  const atlasMarker = '"Topology"';
  const atlasChunks = distFiles.filter(
    (file) =>
      file !== 'index.js' &&
      readFileSync(join(packedDirectory, 'dist', file), 'utf8').includes(atlasMarker),
  );
  if (atlasChunks.length === 0) {
    throw new Error('Packed package is missing the lazily loaded region-map atlas chunk');
  }
  if (readFileSync(join(packedDirectory, 'dist', 'index.js'), 'utf8').includes(atlasMarker)) {
    throw new Error('Region-map atlas data leaked into the eagerly loaded entry bundle');
  }

  const distributableSource = readdirSync(join(packedDirectory, 'dist'))
    .filter((file) => file.endsWith('.js') || file.endsWith('.d.ts'))
    .map((file) => readFileSync(join(packedDirectory, 'dist', file), 'utf8'))
    .join('\n');
  const unresolvedBundledImports = [
    '@extend-ai/react-pptx-model',
    '@extend-ai/react-pptx-wasm',
  ].filter((packageName) => distributableSource.includes(packageName));

  if (unresolvedBundledImports.length > 0) {
    throw new Error(
      `Packed output still imports private workspace packages: ${unresolvedBundledImports.join(', ')}`,
    );
  }

  const declaredPackages = new Set([
    ...Object.keys(manifest.dependencies ?? {}),
    ...Object.keys(manifest.optionalDependencies ?? {}),
    ...Object.keys(manifest.peerDependencies ?? {}),
  ]);
  const importedSpecifiers = [
    ...distributableSource.matchAll(
      /\b(?:from\s*|import\s*(?:\(\s*)?|require\s*\(\s*)["']([^"']+)["']/g,
    ),
  ].map((match) => match[1]);
  const packageName = (specifier) =>
    specifier.startsWith('@')
      ? specifier.split('/').slice(0, 2).join('/')
      : specifier.split('/')[0];
  // Module specifiers never contain whitespace, quotes, or parentheses; this
  // keeps bundled runtime strings such as regl's error messages
  // (`" called from " + callSite`) from tripping the scanner.
  const specifierPattern = /^[@a-zA-Z0-9][\w.+-]*(?:\/[\w.+-]+)*$/;
  const undeclaredImports = [
    ...new Set(
      importedSpecifiers
        .filter(
          (specifier) =>
            specifier &&
            specifierPattern.test(specifier) &&
            !specifier.startsWith('.') &&
            !specifier.startsWith('/') &&
            !specifier.startsWith('node:') &&
            !URL.canParse(specifier),
        )
        .map(packageName)
        .filter((name) => !declaredPackages.has(name)),
    ),
  ];
  if (undeclaredImports.length > 0) {
    throw new Error(`Packed output has undeclared bare imports: ${undeclaredImports.join(', ')}`);
  }

  for (const reactVersion of ['18.3.1', '19.1.0']) {
    const consumerDirectory = join(temporaryDirectory, `consumer-react-${reactVersion}`);
    mkdirSync(consumerDirectory);
    writeFileSync(
      join(consumerDirectory, 'package.json'),
      JSON.stringify(
        {
          private: true,
          type: 'module',
          dependencies: {
            [manifest.name]: `file:${tarball}`,
            react: reactVersion,
            'react-dom': reactVersion,
          },
          devDependencies: {
            '@types/react': reactVersion.startsWith('18.') ? '18.3.24' : '19.1.6',
            '@types/react-dom': reactVersion.startsWith('18.') ? '18.3.7' : '19.1.5',
            typescript: '5.8.3',
          },
        },
        null,
        2,
      ),
    );
    writeFileSync(
      join(consumerDirectory, 'consumer.tsx'),
      `import React from 'react';\nimport { ReactPptxViewer, parsePresentation } from '${manifest.name}';\nimport '${manifest.name}/styles.css';\n\nvoid parsePresentation;\nexport const Viewer = () => <ReactPptxViewer source={new Uint8Array()} />;\n`,
    );
    writeFileSync(
      join(consumerDirectory, 'tsconfig.json'),
      JSON.stringify(
        {
          compilerOptions: {
            jsx: 'react-jsx',
            lib: ['DOM', 'ES2022'],
            module: 'NodeNext',
            moduleResolution: 'NodeNext',
            noEmit: true,
            strict: true,
            target: 'ES2022',
          },
          include: ['consumer.tsx'],
        },
        null,
        2,
      ),
    );
    execFileSync(
      'npm',
      ['install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false'],
      { cwd: consumerDirectory, stdio: 'inherit' },
    );
    execFileSync('node', ['--input-type=module', '--eval', `await import('${manifest.name}')`], {
      cwd: consumerDirectory,
      stdio: 'inherit',
    });
    execFileSync('npx', ['tsc', '--project', 'tsconfig.json'], {
      cwd: consumerDirectory,
      stdio: 'inherit',
    });
  }

  const requestedOutput = process.env.REACT_PPTX_PACK_OUTPUT;
  if (requestedOutput) {
    const outputPath = resolve(workspaceRoot, requestedOutput);
    mkdirSync(dirname(outputPath), { recursive: true });
    copyFileSync(tarball, outputPath);
    console.log(`Saved verified tarball to ${outputPath}`);
  }

  console.log(`Verified ${manifest.name}@${manifest.version} (${packResult.size} bytes)`);
} finally {
  rmSync(temporaryDirectory, { recursive: true, force: true });
}
