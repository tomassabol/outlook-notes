/**
 * Label colors use Outlook's own category palette (Fluent UI v9 "colorPalette*" tokens).
 * Inside Outlook the live tokens are used, so labels follow Outlook's light/dark theme exactly;
 * the hex values are light-theme fallbacks: [Background2, Foreground2, BorderActive].
 */
export const PALETTE: Record<string, [bg: string, fg: string, accent: string]> = {
  Red: ['#f1bbbc', '#751d1f', '#d13438'],
  Cranberry: ['#eeacb2', '#6e0811', '#c50f1f'],
  Pumpkin: ['#efc4ad', '#712d09', '#ca5010'],
  Peach: ['#ffddb3', '#8f4e00', '#ff8c00'],
  Marigold: ['#f9e2ae', '#835b00', '#eaa300'],
  Gold: ['#ecdfa5', '#6c5700', '#c19c00'],
  Brass: ['#e0cea2', '#553e06', '#986f0b'],
  Brown: ['#ddc3b0', '#50301a', '#8e562e'],
  Forest: ['#bdd99b', '#294903', '#498205'],
  LightGreen: ['#a7e3a5', '#0b5a08', '#13a10e'],
  Green: ['#9fd89f', '#094509', '#107c10'],
  Seafoam: ['#a8f0cd', '#00723b', '#00cc6a'],
  LightTeal: ['#a6e9ed', '#00666d', '#00b7c3'],
  Teal: ['#9bd9db', '#02494c', '#038387'],
  Steel: ['#94c8d4', '#00333f', '#005b70'],
  Blue: ['#a9d3f2', '#004377', '#0078d4'],
  RoyalBlue: ['#9abfdc', '#002c4e', '#004e8c'],
  Cornflower: ['#c8d1fa', '#2c3c85', '#4f6bed'],
  Lavender: ['#d2ccf8', '#3f3682', '#7160e8'],
  Purple: ['#c6b1de', '#341a51', '#5c2e91'],
  Grape: ['#d9a7e0', '#4c0d55', '#881798'],
  Pink: ['#f7c0e3', '#80215d', '#e43ba6'],
  Magenta: ['#eca5d1', '#6b0043', '#bf0077'],
  Platinum: ['#cdd6d8', '#3b4447', '#69797e'],
  Anchor: ['#bcc3c7', '#202427', '#394146'],
};

export const PALETTE_NAMES = Object.keys(PALETTE);

/** "LightGreen" → "Light green" */
export function paletteTitle(name: string): string {
  const s = name.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase();
  return s[0].toUpperCase() + s.slice(1);
}

function rgb(hex: string): [number, number, number] {
  const n = parseInt(hex.replace('#', ''), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

/** Accepts a palette name, or a legacy #rrggbb (mapped to the nearest palette color). */
export function paletteName(color: string): string {
  if (color in PALETTE) return color;
  if (!/^#[0-9a-f]{6}$/i.test(color)) return 'Platinum';
  const [r, g, b] = rgb(color);
  let best = 'Platinum';
  let bestD = Infinity;
  for (const [name, [, , accent]] of Object.entries(PALETTE)) {
    const [r2, g2, b2] = rgb(accent);
    const d = (r - r2) ** 2 + (g - g2) ** 2 + (b - b2) ** 2;
    if (d < bestD) {
      bestD = d;
      best = name;
    }
  }
  return best;
}

/** CSS custom properties for a pill/swatch of this color. */
export function paletteVars(color: string): Record<string, string> {
  const name = paletteName(color);
  const [bg, fg, accent] = PALETTE[name];
  return {
    '--pal-bg': `var(--colorPalette${name}Background2, ${bg})`,
    '--pal-fg': `var(--colorPalette${name}Foreground2, ${fg})`,
    '--pal-accent': `var(--colorPalette${name}BorderActive, ${accent})`,
  };
}

export function applyPaletteVars(el: HTMLElement, color: string): void {
  for (const [k, v] of Object.entries(paletteVars(color))) el.style.setProperty(k, v);
}

/** Outlook token names the editor iframe needs copied over (it can't inherit them). */
export function paletteTokenNames(): string[] {
  return PALETTE_NAMES.flatMap((n) => ['Background2', 'Foreground2', 'BorderActive'].map((s) => `--colorPalette${n}${s}`));
}
