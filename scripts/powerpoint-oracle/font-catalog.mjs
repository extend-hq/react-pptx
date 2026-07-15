import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readdir, stat } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const FONT_EXTENSIONS = new Set(['.ttf', '.otf', '.ttc']);

export function defaultOfficeFontRoots() {
  return [
    path.join(os.homedir(), 'Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts'),
    '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts',
  ];
}

async function listFontFiles(root) {
  const files = [];
  const visit = async (directory) => {
    let entries;
    try {
      entries = await readdir(directory, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const absolute = path.join(directory, entry.name);
      if (entry.isDirectory()) await visit(absolute);
      else if (entry.isFile() && FONT_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
        files.push(absolute);
      }
    }
  };
  await visit(root);
  return files;
}

export function normalizeFontKey(value) {
  return value
    .replace(/^[A-Z]{6}\+/, '')
    .replace(
      /(?:-|_)?(?:bolditalic|boldoblique|semibolditalic|semibold|demibold|regular|italic|oblique|bold)$/i,
      '',
    )
    .replace(/(?:mt|ps)$/i, '')
    .replace(/[^a-z0-9]/gi, '')
    .toLowerCase();
}

function parseStyle(style) {
  const lower = style.toLowerCase();
  let weight = 400;
  if (/thin/.test(lower)) weight = 100;
  else if (/extra.?light|ultra.?light/.test(lower)) weight = 200;
  else if (/light/.test(lower)) weight = 300;
  else if (/medium/.test(lower)) weight = 500;
  else if (/semi.?bold|demi.?bold/.test(lower)) weight = 600;
  else if (/extra.?bold|ultra.?bold/.test(lower)) weight = 800;
  else if (/black|heavy/.test(lower)) weight = 900;
  else if (/bold/.test(lower)) weight = 700;
  return { weight, style: /italic|oblique/.test(lower) ? 'italic' : 'normal' };
}

export async function buildOfficeFontCatalog(roots = defaultOfficeFontRoots()) {
  const files = (await Promise.all(roots.map(listFontFiles))).flat().sort();
  const fingerprint = createHash('sha256');
  for (const file of files) {
    const info = await stat(file);
    fingerprint.update(`${file}\0${info.size}\0${Math.trunc(info.mtimeMs)}\n`);
  }

  if (files.length === 0) return { entries: [], fingerprint: fingerprint.digest('hex') };
  const scan = spawnSync('fc-scan', ['--format', '%{file}\t%{family}\t%{style[0]}\n', ...files], {
    encoding: 'utf8',
    maxBuffer: 32 * 1024 * 1024,
  });
  if (scan.status !== 0) {
    throw new Error(`fc-scan failed: ${scan.stderr || `exit ${scan.status}`}`);
  }

  const entries = scan.stdout
    .split('\n')
    .map((line) => line.split('\t'))
    .filter((parts) => parts.length >= 3 && parts[0] && parts[1])
    .flatMap(([file, families, style]) =>
      families.split(',').map((family) => ({
        file,
        family,
        familyKey: normalizeFontKey(family),
        fontStyleName: style,
        ...parseStyle(style),
      })),
    );
  return { entries, fingerprint: fingerprint.digest('hex') };
}

export function matchOracleFonts(oracleFonts, catalogEntries) {
  const requestedKeys = new Set(oracleFonts.map(normalizeFontKey).filter(Boolean));
  const matchedFamilies = new Set();
  for (const requested of requestedKeys) {
    const exact = catalogEntries.find((entry) => entry.familyKey === requested);
    if (exact) {
      matchedFamilies.add(exact.familyKey);
      continue;
    }
    const fuzzy = catalogEntries
      .filter(
        (entry) =>
          entry.familyKey.length >= 4 &&
          (requested.includes(entry.familyKey) || entry.familyKey.includes(requested)),
      )
      .sort((a, b) => b.familyKey.length - a.familyKey.length)[0];
    if (fuzzy) matchedFamilies.add(fuzzy.familyKey);
  }
  return catalogEntries.filter((entry) => matchedFamilies.has(entry.familyKey));
}
