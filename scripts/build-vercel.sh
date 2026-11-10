#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
build_tools="$repo_root/node_modules/.cache/vercel-tools"
export PATH="$build_tools/bin:${CARGO_HOME:-$HOME/.cargo}/bin:$PATH"

if ! command -v rustup >/dev/null 2>&1; then
  curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs |
    sh -s -- -y --profile minimal --default-toolchain none --no-modify-path
fi

export RUSTUP_TOOLCHAIN=1.93.1
rustup toolchain install "$RUSTUP_TOOLCHAIN" --profile minimal --target wasm32-unknown-unknown
cargo install wasm-pack --version 0.13.1 --locked --root "$build_tools"

pnpm --dir "$repo_root" build
