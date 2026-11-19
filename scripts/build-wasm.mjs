import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const upstream = process.env.RUST_OFFICE_PATH;
let crate = root;
if (upstream) {
  const dukeCrate = resolve(upstream, 'crates/duke-slides-wasm');
  crate = existsSync(dukeCrate) ? dukeCrate : resolve(upstream, 'crates/rust-pptx-wasm');
}
const result = spawnSync(
  'wasm-pack',
  [
    'build',
    crate,
    '--target',
    'web',
    '--out-name',
    'pptx_wasm',
    '--out-dir',
    resolve(root, 'packages/wasm/pkg'),
    '--release',
    '--',
    '--locked',
  ],
  { cwd: root, stdio: 'inherit' },
);
if (result.error) throw result.error;
process.exit(result.status ?? 1);
