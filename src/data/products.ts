import type { Color, ReferenceMap } from '../types';

export const COLORS: Color[] = [
  /* ── Basiques ── */
  { n: 'Blanc',              h: '#FFFFFF' },
  { n: 'Noir',               h: '#1A1A1A' },
  { n: 'Black',              h: '#0D0D0D' },
  { n: 'Gris',               h: '#9B9B9B' },
  { n: 'Navy',               h: '#001F4D' },
  { n: 'Navy Blue',          h: '#003087' },
  { n: 'Sport Navy',         h: '#1F3A5F' },
  { n: 'Royal Blue',         h: '#0A3D91' },
  { n: 'True Royal',         h: '#2563EB' },
  { n: 'Royal',              h: '#1155CC' },
  { n: 'Rouge',              h: '#CC1517' },
  { n: 'Orange',             h: '#FF6600' },
  /* ── Gris ── */
  { n: 'Dark Grey',          h: '#3D3D3D' },
  { n: 'Oxford Grey',        h: '#545454' },
  { n: 'Iron Grey',          h: '#666666' },
  { n: 'Heather Dark Grey',  h: '#4A4A4A' },
  { n: 'Fine Grey',          h: '#AEAEB2' },
  /* ── Bleus ── */
  { n: 'Sky Blue',           h: '#87CEEB' },
  { n: 'Light Blue',         h: '#ADD8E6' },
  { n: 'Cool Blue',          h: '#4A90D9' },
  { n: 'Light Royal Blue',   h: '#4169E1' },
  { n: 'Baltic Blue',        h: '#2E5FA3' },
  { n: 'Sea Blue',           h: '#2C6E8A' },
  { n: 'Tropical Blue',      h: '#0077B6' },
  { n: 'Aqua Blue',          h: '#00B5CC' },
  { n: 'Aquamarine',         h: '#7FFFD4' },
  { n: 'Vintage Blue',       h: '#4A6D8C' },
  /* ── Kakis / Olives ── */
  { n: 'Kaki',               h: '#8B8456' },
  { n: 'Khaki',              h: '#C3B091' },
  { n: 'Dark Kaki',          h: '#5C5A3A' },
  { n: 'Organic Kaki',       h: '#9C8E5A' },
  { n: 'Urban Kaki',         h: '#6E6A3D' },
  { n: 'Olive',              h: '#808000' },
  { n: 'Olive Green',        h: '#6B7C43' },
  { n: 'Light Olive Green',  h: '#A8BE8C' },
  { n: 'Light Army Green',   h: '#7D8A5C' },
  /* ── Verts ── */
  { n: 'Forest Green',       h: '#228B22' },
  { n: 'Green Field',        h: '#5A8A3C' },
  { n: 'Jade Green',         h: '#00A86B' },
  { n: 'Almond Green',       h: '#8DB48E' },
  { n: 'Mint Green',         h: '#98FF98' },
  { n: 'Bright Seafoam',     h: '#3EFFC5' },
  /* ── Sables / Naturels ── */
  { n: 'Sand',               h: '#C2A679' },
  { n: 'Light Sand',         h: '#EAD9B8' },
  { n: 'West Sand',          h: '#BCA98A' },
  { n: 'Wet Sand',           h: '#B0A08A' },
  { n: 'Union Beige',        h: '#D4C5A9' },
  { n: 'Driftwood',          h: '#B5A380' },
  { n: 'Raw Natural',        h: '#D9CBAD' },
  { n: 'Ivory',              h: '#FFFFF0' },
  { n: 'Ivoir',              h: '#F5EDD8' },
  { n: 'Coconut Milk',       h: '#FAF0E6' },
  { n: 'Washed Cream Coffee', h: '#C4A97A' },
  /* ── Coraux / Rouges chauds ── */
  { n: 'Coral',              h: '#FF7F6E' },
  { n: 'Soft Coral',         h: '#F4845F' },
  { n: 'Terracotta Red',     h: '#C25B3C' },
  { n: 'Paprika',            h: '#A31D1D' },
  { n: 'Clementine Heather', h: '#D4603A' },
  { n: 'Washed Sienna',      h: '#A0522D' },
  /* ── Pastel / Doux ── */
  { n: 'Peach',              h: '#FFCBA4' },
  { n: 'Petal Rose',         h: '#F4A7B9' },
  { n: 'Pineapple',          h: '#E8C84A' },
];

export const LOGO_COLORS: Color[] = [
  { n: 'Noir',   h: '#1A1A1A' },
  { n: 'Blanc',  h: '#FFFFFF' },
  { n: 'Marine', h: '#00205B' },
  { n: 'Rouge',  h: '#FF3B30' },
  { n: 'Or',     h: '#C5A44E' },
  { n: 'Argent', h: '#A8A8A8' },
  { n: 'Royal',  h: '#003087' },
  { n: 'Vert',   h: '#228B22' },
  { n: 'Jaune',  h: '#FFD700' },
  { n: 'Rose',   h: '#FF69B4' },
  { n: 'Orange', h: '#FF6600' },
  { n: 'Violet', h: '#6B46C1' },
];

export const REFERENCES: ReferenceMap = {
  'H-001': { fournisseur: 'NS300', prix: 35, bio: true,  largeurs: { XS: 220, S: 250, M: 270, L: 300, XL: 320, XXL: 320 } },
  'H-002': { fournisseur: 'NS300', prix: 39, bio: false, largeurs: { XS: 220, S: 250, M: 270, L: 300, XL: 320, XXL: 320 } },
};

export const COLLECTIONS = ['Homme', 'Femme', 'Enfant', 'Accessoire'] as const;
export const SIZES = ['XS', 'S', 'M', 'L', 'XL', 'XXL'] as const;
export const PRICE_TSHIRT  = [15, 20, 25, 30, 35, 39, 40, 45, 50];
export const PRICE_PERSO   = [0, 5, 10, 15, 20, 25, 30];

// Logo SVG data
export const LOGOS_BEA = [
  {
    id: 'bea-1', n: 'Classic',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><circle cx="60" cy="60" r="50" stroke="currentColor" stroke-width="3" fill="none"/><text x="60" y="52" text-anchor="middle" font-size="28" font-weight="800" fill="currentColor" font-family="sans-serif">BEA</text><text x="60" y="74" text-anchor="middle" font-size="12" font-weight="600" fill="currentColor" letter-spacing="4" font-family="sans-serif">ATELIER</text><line x1="28" y1="82" x2="92" y2="82" stroke="currentColor" stroke-width="1.5"/><text x="60" y="98" text-anchor="middle" font-size="10" fill="currentColor" font-family="sans-serif">— 16 —</text></svg>`,
  },
  {
    id: 'bea-2', n: 'Minimal',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><rect x="15" y="25" width="90" height="70" rx="4" stroke="currentColor" stroke-width="2.5" fill="none"/><text x="60" y="60" text-anchor="middle" font-size="32" font-weight="900" fill="currentColor" font-family="sans-serif">B</text><text x="60" y="82" text-anchor="middle" font-size="9" font-weight="600" fill="currentColor" letter-spacing="6" font-family="sans-serif">BEAUTE</text></svg>`,
  },
  {
    id: 'bea-3', n: 'Script',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><text x="60" y="55" text-anchor="middle" font-family="Georgia,serif" font-size="36" font-style="italic" font-weight="700" fill="currentColor">Bea</text><line x1="20" y1="65" x2="100" y2="65" stroke="currentColor" stroke-width="1"/><text x="60" y="82" text-anchor="middle" font-size="8" font-weight="600" fill="currentColor" letter-spacing="5" font-family="sans-serif">COLLECTION 16</text></svg>`,
  },
];

export const LOGOS_SXM = [
  {
    id: 'sxm-1', n: 'Island',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><path d="M60 15 L95 45 L85 95 L35 95 L25 45 Z" stroke="currentColor" stroke-width="2.5" fill="none"/><text x="60" y="58" text-anchor="middle" font-size="22" font-weight="900" fill="currentColor" font-family="sans-serif">SXM</text><text x="60" y="76" text-anchor="middle" font-size="9" font-weight="500" fill="currentColor" letter-spacing="3" font-family="sans-serif">ISLAND</text><text x="60" y="90" text-anchor="middle" font-size="8" fill="currentColor" font-family="sans-serif">— 12 —</text></svg>`,
  },
  {
    id: 'sxm-2', n: 'Wave',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><text x="60" y="50" text-anchor="middle" font-size="30" font-weight="900" fill="currentColor" font-family="sans-serif">SXM</text><path d="M20 65 Q35 55 50 65 Q65 75 80 65 Q95 55 100 65" stroke="currentColor" stroke-width="2.5" fill="none"/><text x="60" y="100" text-anchor="middle" font-size="8" font-weight="600" fill="currentColor" letter-spacing="4" font-family="sans-serif">SAINT-MARTIN</text></svg>`,
  },
  {
    id: 'sxm-3', n: 'Geo',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><circle cx="60" cy="55" r="35" stroke="currentColor" stroke-width="2" fill="none"/><circle cx="60" cy="55" r="25" stroke="currentColor" stroke-width="1.5" fill="none" opacity="0.4"/><text x="60" y="62" text-anchor="middle" font-size="18" font-weight="800" fill="currentColor" font-family="sans-serif">SXM</text><text x="60" y="105" text-anchor="middle" font-size="8" font-weight="600" fill="currentColor" letter-spacing="3" font-family="sans-serif">12 — CARAIBES</text></svg>`,
  },
];

export const LOGOS_COR = [
  {
    id: 'cor-1', n: 'Shield',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><path d="M60 18 L95 35 L95 70 Q95 95 60 108 Q25 95 25 70 L25 35 Z" stroke="currentColor" stroke-width="2.5" fill="none"/><text x="60" y="62" text-anchor="middle" font-size="20" font-weight="900" fill="currentColor" font-family="sans-serif">COR</text><text x="60" y="80" text-anchor="middle" font-size="9" font-weight="500" fill="currentColor" letter-spacing="2" font-family="sans-serif">— 01 —</text></svg>`,
  },
  {
    id: 'cor-2', n: 'Typo',
    s: `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><text x="60" y="45" text-anchor="middle" font-size="40" font-weight="900" fill="currentColor" font-family="sans-serif">C</text><line x1="25" y1="55" x2="95" y2="55" stroke="currentColor" stroke-width="2"/><text x="60" y="72" text-anchor="middle" font-size="11" font-weight="600" fill="currentColor" letter-spacing="8" font-family="sans-serif">CORSICA</text><text x="60" y="92" text-anchor="middle" font-size="9" font-weight="500" fill="currentColor" font-family="sans-serif">Edition 01</text></svg>`,
  },
];

export const LOGO_SECTIONS = [
  { key: 'bea', label: 'BEA Atelier', logos: LOGOS_BEA },
  { key: 'sxm', label: 'SXM Island',  logos: LOGOS_SXM },
  { key: 'cor', label: 'COR Corsica', logos: LOGOS_COR },
];

export const FABRIC_RE = /rgb\(100%,\s*100%,\s*100%\)/gi;

/** Replace the white fabric color in the SVG with the selected hex */
export function applyFabricColor(svg: string, hex: string): string {
  return svg.replace(FABRIC_RE, hex);
}

/** Returns true if the color is light (to show dark border on swatch) */
export function isLightColor(hex: string): boolean {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return (r * 299 + g * 587 + b * 114) / 1000 > 200;
}

/** Minimal monochrome product type icons */
export const PRODUCT_ICONS = {
  'tshirt': `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><path d="M 30 20 L 45 35 L 45 90 Q 45 100 55 100 Q 65 100 65 90 L 65 35 L 80 20 L 80 30 L 60 50 L 60 100 L 50 100 L 50 50 L 30 30 Z" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linejoin="round" stroke-linecap="round"/></svg>`,
  'hoodie': `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><path d="M 40 35 Q 30 35 25 45 L 25 100 Q 25 105 30 105 L 90 105 Q 95 105 95 100 L 95 45 Q 90 35 80 35 L 70 35 L 70 20 Q 70 15 65 15 L 55 15 Q 50 15 50 20 L 50 35 L 40 35 M 60 35 L 60 50 Q 60 55 55 55 L 55 75 L 65 75 L 65 55 Q 60 55 60 50 Z" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linejoin="round" stroke-linecap="round"/></svg>`,
  'mug': `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><rect x="25" y="30" width="55" height="60" rx="4" stroke="currentColor" stroke-width="2.5" fill="none"/><path d="M 80 40 Q 95 45 95 60 Q 95 75 80 80" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linecap="round"/><line x1="35" y1="50" x2="70" y2="50" stroke="currentColor" stroke-width="1.5" opacity="0.5"/></svg>`,
  'hat': `<svg viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg"><path d="M 30 60 L 35 35 L 85 35 L 90 60 Z" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linejoin="round"/><path d="M 28 62 L 92 62 Q 95 62 95 65 L 95 90 Q 95 100 85 105 L 35 105 Q 25 100 25 90 L 25 65 Q 25 62 28 62" stroke="currentColor" stroke-width="2.5" fill="none" stroke-linejoin="round"/></svg>`,
};
