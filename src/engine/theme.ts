// Theme colors/fonts and DrawingML color resolution.
import { attr, child, children, num, path } from './xml';

export interface Theme {
  colors: Record<string, string>; // dk1, lt1, accent1 ... as RRGGBB
  majorLatin: string;
  minorLatin: string;
  majorEa: string;
  minorEa: string;
  lineWidths: number[]; // EMU widths of lnStyleLst entries
}

export const DEFAULT_THEME: Theme = {
  colors: {
    dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
    accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000',
    accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
  },
  majorLatin: 'Calibri Light', minorLatin: 'Calibri', majorEa: '', minorEa: '',
  lineWidths: [9525, 25400, 38100],
};

export function parseTheme(doc: XMLDocument | null): Theme {
  if (!doc) return DEFAULT_THEME;
  const te = path(doc.documentElement, 'a:themeElements');
  const colors: Record<string, string> = { ...DEFAULT_THEME.colors };
  for (const c of children(child(te, 'a:clrScheme'))) {
    const v = child(c, 'a:srgbClr')?.getAttribute('val') ?? child(c, 'a:sysClr')?.getAttribute('lastClr');
    if (v) colors[c.localName] = v.toUpperCase();
  }
  const fs = child(te, 'a:fontScheme');
  const font = (kind: string, sub: string) => attr(path(fs, kind, sub), 'typeface') || '';
  const lineWidths = children(path(te, 'a:fmtScheme', 'a:lnStyleLst'), 'a:ln').map((l) => num(l, 'w', 9525));
  return {
    colors,
    majorLatin: font('a:majorFont', 'a:latin') || DEFAULT_THEME.majorLatin,
    minorLatin: font('a:minorFont', 'a:latin') || DEFAULT_THEME.minorLatin,
    majorEa: font('a:majorFont', 'a:ea'),
    minorEa: font('a:minorFont', 'a:ea'),
    lineWidths: lineWidths.length ? lineWidths : DEFAULT_THEME.lineWidths,
  };
}

export interface ColorContext {
  theme: Theme;
  clrMap: Record<string, string>; // bg1 -> lt1 ...
  phClr?: string | null; // style placeholder color as RRGGBB
}

export const DEFAULT_CLRMAP: Record<string, string> = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4',
  accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink',
};

export function parseClrMap(el: Element | null): Record<string, string> {
  if (!el) return DEFAULT_CLRMAP;
  const m: Record<string, string> = { ...DEFAULT_CLRMAP };
  for (const a of Array.from(el.attributes)) m[a.name] = a.value;
  return m;
}

const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080', darkGray: 'A9A9A9', lightGray: 'D3D3D3',
};

interface RGBA { r: number; g: number; b: number; a: number }

function hexToRgb(h: string): RGBA {
  const n = parseInt(h.slice(0, 6), 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

function rgbToHsl({ r, g, b }: RGBA): [number, number, number] {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number, a: number): RGBA {
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255, a };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return { r: f(h + 1 / 3) * 255, g: f(h) * 255, b: f(h - 1 / 3) * 255, a };
}

function applyMods(c: RGBA, el: Element): RGBA {
  for (const m of children(el)) {
    const v = num(m, 'val') / 100000;
    switch (m.localName) {
      case 'alpha': c.a = v; break;
      case 'lumMod': case 'lumOff': {
        const [h, s, l] = rgbToHsl(c);
        const nl = m.localName === 'lumMod' ? l * v : l + v;
        c = hslToRgb(h, s, Math.min(1, Math.max(0, nl)), c.a);
        break;
      }
      case 'satMod': {
        const [h, s, l] = rgbToHsl(c);
        c = hslToRgb(h, Math.min(1, s * v), l, c.a);
        break;
      }
      case 'tint': c = { r: c.r + (255 - c.r) * (1 - v), g: c.g + (255 - c.g) * (1 - v), b: c.b + (255 - c.b) * (1 - v), a: c.a }; break;
      case 'shade': c = { r: c.r * v, g: c.g * v, b: c.b * v, a: c.a }; break;
    }
  }
  return c;
}

/** Resolve the color element (a:srgbClr etc.) inside container to an RGBA. */
function resolveRgba(container: Element | null, ctx: ColorContext): RGBA | null {
  if (!container) return null;
  for (const c of children(container)) {
    let hex: string | null = null;
    switch (c.localName) {
      case 'srgbClr': hex = c.getAttribute('val'); break;
      case 'sysClr': hex = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000'); break;
      case 'prstClr': hex = PRESET[c.getAttribute('val') || ''] || '000000'; break;
      case 'scrgbClr': {
        const p = (k: string) => Math.round((num(c, k) / 100000) * 255).toString(16).padStart(2, '0');
        hex = p('r') + p('g') + p('b');
        break;
      }
      case 'schemeClr': {
        const v = c.getAttribute('val') || '';
        if (v === 'phClr') hex = ctx.phClr ?? '000000';
        else hex = ctx.theme.colors[ctx.clrMap[v] ?? v] ?? ctx.theme.colors[v] ?? '000000';
        break;
      }
      case 'hslClr': {
        const rgb = hslToRgb(num(c, 'hue') / 21600000, num(c, 'sat') / 100000, num(c, 'lum') / 100000, 1);
        return applyMods(rgb, c);
      }
      default: continue;
    }
    if (hex) return applyMods(hexToRgb(hex), c);
  }
  return null;
}

export function resolveColor(container: Element | null, ctx: ColorContext): string | null {
  const c = resolveRgba(container, ctx);
  if (!c) return null;
  const r = Math.round(c.r), g = Math.round(c.g), b = Math.round(c.b);
  return c.a >= 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${c.a.toFixed(3)})`;
}

export function resolveHex(container: Element | null, ctx: ColorContext): string | null {
  const c = resolveRgba(container, ctx);
  if (!c) return null;
  return [c.r, c.g, c.b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

export function cssToHex(css: string): string | null {
  const m = css.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (m) return [m[1], m[2], m[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase();
  const h = css.match(/^#([0-9a-f]{6})$/i);
  return h ? h[1].toUpperCase() : null;
}
