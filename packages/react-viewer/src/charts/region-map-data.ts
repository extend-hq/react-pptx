/**
 * Region-map registry shared by the chart renderer and the lazily loaded
 * atlas chunk. The multi-megabyte TopoJSON boundary data only loads when a
 * presentation actually contains an Excel map chart; until then the renderer
 * sees an empty registry and draws the map without shapes.
 */
import type { Feature, Geometry } from 'geojson';

export type RegionMapFeature = Feature<
  Geometry,
  { name?: string; regionSet?: 'country' | 'us-state'; stateCode?: string }
>;

export interface RegionMapAtlas {
  worldCountryFeatures: RegionMapFeature[];
  usStateFeatures: RegionMapFeature[];
  countryFeaturesByKey: Map<string, RegionMapFeature>;
  usStateFeaturesByKey: Map<string, RegionMapFeature>;
}

export function normalizeRegionMapKey(value: unknown): string {
  if (typeof value !== 'string') {
    return '';
  }
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/&/g, ' and ')
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

export const REGION_MAP_COUNTRY_ALIASES = new Map<string, string>([
  ['us', 'united states of america'],
  ['usa', 'united states of america'],
  ['u s a', 'united states of america'],
  ['united states', 'united states of america'],
  ['united states america', 'united states of america'],
  ['u s', 'united states of america'],
  ['uk', 'united kingdom'],
  ['u k', 'united kingdom'],
  ['uae', 'united arab emirates'],
  ['u a e', 'united arab emirates'],
  ['south korea', 'korea, south'],
  ['north korea', 'korea, north'],
  ['russia', 'russian federation'],
  ['vietnam', 'viet nam'],
  ['czech republic', 'czechia'],
  ['ivory coast', "cote d'ivoire"],
  ['côte divoire', "cote d'ivoire"],
]);

export const REGION_MAP_US_STATE_ALIASES = new Map<string, string>([
  ['district of columbia', 'district of columbia'],
  ['washington dc', 'district of columbia'],
  ['washington d c', 'district of columbia'],
  ['dc', 'district of columbia'],
  ['d c', 'district of columbia'],
]);

let loadedAtlas: RegionMapAtlas | null = null;
let pendingAtlas: Promise<RegionMapAtlas> | null = null;

/** Synchronous accessor used by the renderer; null until the chunk loads. */
export function getRegionMapAtlas(): RegionMapAtlas | null {
  return loadedAtlas;
}

/** Loads the atlas chunk on demand; resolves immediately once cached. */
export function loadRegionMapAtlas(): Promise<RegionMapAtlas> {
  if (loadedAtlas) return Promise.resolve(loadedAtlas);
  pendingAtlas ??= import('./region-map-atlas').then((module) => {
    loadedAtlas = module.REGION_MAP_ATLAS;
    return loadedAtlas;
  });
  return pendingAtlas;
}
