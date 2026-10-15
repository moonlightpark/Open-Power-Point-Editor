// Text bodies: style inheritance, rendering to editable DOM, and committing DOM edits back to XML.
import { RunProps, SlideContext, findInherited, placeholder, setRunProps } from './document';
import { ColorContext, cssToHex, resolveColor } from './theme';
import { NS, attr, child, children, create, num, path } from './xml';

export interface TextScope {
  ctx: SlideContext;
  color: ColorContext;
  /** lstStyle elements (most specific first) — shape, layout ph, master ph. */
  lstStyles: Element[];
  /** bodyPr elements (most specific first). */
  bodyPrs: Element[];
  /** master txStyles element for this text (titleStyle/bodyStyle/otherStyle). */
  masterStyle: Element | null;
  defaultStyle: Element | null; // presentation defaultTextStyle
  fontRefColor: string | null; // from p:style/a:fontRef
  fontRefFont: 'major' | 'minor' | null;
}

const TITLE_TYPES = new Set(['title', 'ctrTitle']);
const BODY_TYPES = new Set(['body', 'subTitle', 'obj']);

/** Build the inheritance scope for a shape's (or table cell's) text. */
export function textScope(ctx: SlideContext, pres: XMLDocument, shape: Element, txBody: Element, owner: 'slide' | 'layout' | 'master' = 'slide', inTable = false): TextScope {
  const lstStyles: Element[] = [];
  const bodyPrs: Element[] = [];
  const push = (tb: Element | null) => {
    const l = child(tb, 'a:lstStyle');
    const b = child(tb, 'a:bodyPr');
    if (l) lstStyles.push(l);
    if (b) bodyPrs.push(b);
  };
  push(txBody);
  const ph = placeholder(shape);
  let masterStyle: Element | null = null;
  const txStyles = path(ctx.master?.documentElement, 'p:txStyles');
  if (ph && !inTable) {
    if (owner === 'slide') {
      const lay = findInherited(ph, ctx.layout, false);
      if (lay) push(child(lay, 'p:txBody'));
      const mph = lay ? placeholder(lay)! : ph;
      const mas = findInherited(mph, ctx.master, true);
      if (mas) push(child(mas, 'p:txBody'));
    } else if (owner === 'layout') {
      const mas = findInherited(ph, ctx.master, true);
      if (mas) push(child(mas, 'p:txBody'));
    }
    const t = ph.getAttribute('type') || 'obj';
    if (TITLE_TYPES.has(t)) masterStyle = child(txStyles, 'p:titleStyle');
    else if (BODY_TYPES.has(t) || !ph.getAttribute('type')) masterStyle = child(txStyles, 'p:bodyStyle');
    else masterStyle = child(txStyles, 'p:otherStyle');
  } else if (inTable) {
    masterStyle = child(txStyles, 'p:otherStyle');
  }
  const fontRef = path(shape, 'p:style', 'a:fontRef');
  const color: ColorContext = { ...ctx.color };
  const fontRefColor = fontRef && !inTable ? resolveColor(fontRef, color) : null;
  const idx = fontRef?.getAttribute('idx');
  return {
    ctx, color, lstStyles, bodyPrs, masterStyle,
    defaultStyle: child(pres.documentElement, 'p:defaultTextStyle'),
    fontRefColor,
    fontRefFont: idx === 'major' ? 'major' : idx === 'minor' ? 'minor' : null,
  };
}

/** Shape/placeholder list-style levels (own) followed by master/presentation defaults (base). */
function lvlParts(scope: TextScope, lvl: number): { own: Element[]; base: Element[] } {
  const tag = `a:lvl${lvl + 1}pPr`;
  const own: Element[] = [];
  for (const l of scope.lstStyles) { const e = child(l, tag); if (e) own.push(e); }
  const base = [child(scope.masterStyle, tag), child(scope.defaultStyle, tag)].filter((e): e is Element => !!e);
  return { own, base };
}

function lvlChain(scope: TextScope, lvl: number): Element[] {
  const { own, base } = lvlParts(scope, lvl);
  return [...own, ...base];
}

function bodyAttr(scope: TextScope, name: string): string | null {
  for (const b of scope.bodyPrs) { const v = b.getAttribute(name); if (v != null) return v; }
  return null;
}

function bodyChild(scope: TextScope, name: string): Element | null {
  for (const b of scope.bodyPrs) { const v = child(b, name); if (v) return v; }
  return null;
}

export interface BodyLayout {
  insets: [number, number, number, number]; // l t r b (EMU)
  anchor: 'top' | 'middle' | 'bottom';
  wrap: boolean;
  vertical: string | null;
  autofit: 'none' | 'shape' | 'norm';
  fontScale: number;
  lnSpcReduction: number;
}

export function bodyLayout(scope: TextScope): BodyLayout {
  const ins = (n: string, d: number) => { const v = bodyAttr(scope, n); return v == null ? d : Number(v); };
  const a = bodyAttr(scope, 'anchor');
  const own = scope.bodyPrs[0];
  const norm = child(own, 'a:normAutofit');
  return {
    insets: [ins('lIns', 91440), ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720)],
    anchor: a === 'ctr' ? 'middle' : a === 'b' ? 'bottom' : 'top',
    wrap: bodyAttr(scope, 'wrap') !== 'none',
    vertical: bodyAttr(scope, 'vert'),
    autofit: bodyChild(scope, 'a:spAutoFit') ? 'shape' : norm ? 'norm' : 'none',
    fontScale: norm ? num(norm, 'fontScale', 100000) / 100000 : 1,
    lnSpcReduction: norm ? num(norm, 'lnSpcReduction', 0) / 100000 : 0,
  };
}

// ------------------------------------------------------------ effective properties

export interface ParaStyle {
  algn: string;
  marL: number; indent: number; // EMU
  lnSpcPct: number | null; lnSpcPts: number | null;
  spcBefPts: number; spcBefPct: number; spcAftPts: number; spcAftPct: number;
  bullet: { kind: 'none' } | { kind: 'char'; char: string; font: string | null; color: string | null; size: number } | { kind: 'num'; scheme: string; start: number; font: string | null; color: string | null; size: number };
}

function firstAttr(els: (Element | null)[], name: string): string | null {
  for (const e of els) { const v = e?.getAttribute(name); if (v != null) return v; }
  return null;
}

function firstChild(els: (Element | null)[], name: string): Element | null {
  for (const e of els) { const v = child(e, name); if (v) return v; }
  return null;
}

export function paraStyle(scope: TextScope, pPr: Element | null, lvl: number): ParaStyle {
  const chain = [pPr, ...lvlChain(scope, lvl)];
  const spc = (tag: string): [number, number] => {
    const e = firstChild(chain, tag);
    return [num(child(e, 'a:spcPts'), 'val', 0) / 100, num(child(e, 'a:spcPct'), 'val', 0) / 100000];
  };
  const ln = firstChild(chain, 'a:lnSpc');
  const [spcBefPts, spcBefPct] = spc('a:spcBef');
  const [spcAftPts, spcAftPct] = spc('a:spcAft');
  // Bullet: the first level that specifies any bullet type wins.
  let bullet: ParaStyle['bullet'] = { kind: 'none' };
  for (const e of chain) {
    if (!e) continue;
    if (child(e, 'a:buNone')) { bullet = { kind: 'none' }; break; }
    const ch = child(e, 'a:buChar'), an = child(e, 'a:buAutoNum');
    if (ch || an) {
      const font = attr(firstChild(chain, 'a:buFont'), 'typeface');
      const clr = firstChild(chain, 'a:buClr');
      const color = clr ? resolveColor(clr, scope.color) : null;
      const size = num(firstChild(chain, 'a:buSzPct'), 'val', 100000) / 100000;
      bullet = ch
        ? { kind: 'char', char: ch.getAttribute('char') || '•', font, color, size }
        : { kind: 'num', scheme: an!.getAttribute('type') || 'arabicPeriod', start: num(an, 'startAt', 1), font, color, size };
      break;
    }
  }
  return {
    algn: firstAttr(chain, 'algn') || 'l',
    marL: Number(firstAttr(chain, 'marL') ?? 0),
    indent: Number(firstAttr(chain, 'indent') ?? 0),
    lnSpcPct: ln && child(ln, 'a:spcPct') ? num(child(ln, 'a:spcPct'), 'val') / 100000 : null,
    lnSpcPts: ln && child(ln, 'a:spcPts') ? num(child(ln, 'a:spcPts'), 'val') / 100 : null,
    spcBefPts, spcBefPct, spcAftPts, spcAftPct, bullet,
  };
}

export interface RunStyle {
  sz: number; b: boolean; i: boolean; u: boolean; strike: boolean;
  color: string; fonts: string[]; baseline: number; highlight: string | null;
  caps: boolean;
}

function themeFont(scope: TextScope, tf: string | null): string | null {
  if (!tf) return null;
  const t = scope.ctx.theme;
  switch (tf) {
    case '+mj-lt': return t.majorLatin;
    case '+mn-lt': return t.minorLatin;
    case '+mj-ea': return t.majorEa || null;
    case '+mn-ea': return t.minorEa || null;
    default: return tf;
  }
}

export function runStyle(scope: TextScope, rPr: Element | null, lvl: number, pPr: Element | null): RunStyle {
  const parts = lvlParts(scope, lvl);
  const ownChain = [rPr, child(pPr, 'a:defRPr'), ...parts.own.map((e) => child(e, 'a:defRPr'))];
  const chain = [...ownChain, ...parts.base.map((e) => child(e, 'a:defRPr'))];
  const bool = (n: string) => { const v = firstAttr(chain, n); return v === '1' || v === 'true'; };
  let color: string | null = null;
  // A shape style's fontRef color outranks the master/presentation defaults.
  for (const e of scope.fontRefColor ? ownChain : chain) {
    if (!e) continue;
    const f = child(e, 'a:solidFill');
    if (f) { color = resolveColor(f, scope.color); break; }
    if (child(e, 'a:gradFill')) { color = resolveColor(path(e, 'a:gradFill', 'a:gsLst', 'a:gs'), scope.color); break; }
    if (e === rPr && child(e, 'a:noFill')) { color = 'transparent'; break; }
  }
  if (!color) color = scope.fontRefColor ?? resolveColor(parseColorStub('tx1'), scope.color) ?? '#000';
  const latin = themeFont(scope, attr(firstChild(chain, 'a:latin'), 'typeface'));
  const ea = themeFont(scope, attr(firstChild(chain, 'a:ea'), 'typeface'));
  const t = scope.ctx.theme;
  const refLatin = scope.fontRefFont === 'major' ? t.majorLatin : t.minorLatin;
  const refEa = scope.fontRefFont === 'major' ? t.majorEa : t.minorEa;
  const fonts = [latin ?? refLatin, ea ?? refEa].filter((f): f is string => !!f);
  const u = firstAttr(chain, 'u');
  const strike = firstAttr(chain, 'strike');
  const hl = firstChild([rPr], 'a:highlight');
  return {
    sz: Number(firstAttr(chain, 'sz') ?? 1800) / 100,
    b: bool('b'), i: bool('i'),
    u: !!u && u !== 'none',
    strike: !!strike && strike !== 'noStrike',
    color, fonts,
    baseline: Number(firstAttr(chain, 'baseline') ?? 0) / 100000,
    highlight: hl ? resolveColor(hl, scope.color) : null,
    caps: firstAttr(chain, 'cap') === 'all',
  };
}

let stubDoc: XMLDocument | null = null;
function parseColorStub(scheme: string): Element {
  stubDoc ??= document.implementation.createDocument(NS.a, 'a:root', null);
  const holder = stubDoc.createElementNS(NS.a, 'a:solidFill');
  const c = stubDoc.createElementNS(NS.a, 'a:schemeClr');
  c.setAttribute('val', scheme);
  holder.appendChild(c);
  return holder;
}

export const FALLBACK_FONTS = `'Apple SD Gothic Neo','Malgun Gothic','Noto Sans KR',sans-serif`;

export function fontStack(fonts: string[]) {
  return [...new Set(fonts)].map((f) => `'${f.replace(/'/g, '')}'`).concat(FALLBACK_FONTS).join(',');
}

function autoNum(scheme: string, n: number): string {
  const alpha = (k: number) => { let s = ''; k--; do { s = String.fromCharCode(97 + (k % 26)) + s; k = Math.floor(k / 26) - 1; } while (k >= 0); return s; };
  const roman = (k: number) => {
    const map: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let s = ''; for (const [v, r] of map) while (k >= v) { s += r; k -= v; } return s;
  };
  const circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';
  let core: string;
  if (scheme.startsWith('alphaLc')) core = alpha(n);
  else if (scheme.startsWith('alphaUc')) core = alpha(n).toUpperCase();
  else if (scheme.startsWith('romanLc')) core = roman(n);
  else if (scheme.startsWith('romanUc')) core = roman(n).toUpperCase();
  else if (scheme.startsWith('circleNum')) core = circled[n - 1] ?? String(n);
  else core = String(n);
  if (scheme.endsWith('ParenBoth')) return `(${core})`;
  if (scheme.endsWith('ParenR')) return `${core})`;
  if (scheme.endsWith('Period')) return `${core}.`;
  return core;
}

// ------------------------------------------------------------ rendering

export interface RenderedText {
  root: HTMLElement; // positioned body (padding/anchor)
  content: HTMLElement; // editable paragraph container
  paras: Element[]; // original a:p elements, index = data-pi
  layout: BodyLayout;
  isEmpty: boolean;
}

const pt = (emu: number) => emu / 12700;

export function hasText(txBody: Element | null): boolean {
  if (!txBody) return false;
  for (const p of children(txBody, 'a:p')) for (const r of children(p)) if ((r.localName === 'r' || r.localName === 'fld') && (child(r, 'a:t')?.textContent ?? '') !== '') return true;
  return false;
}

/**
 * Render a txBody. `prompt` paragraphs (from a layout placeholder) are shown greyed out
 * when the body is empty and the caller is the editor (not a thumbnail or slide show).
 */
export function renderText(scope: TextScope, txBody: Element, opts: { prompt?: Element | null } = {}): RenderedText {
  const layout = bodyLayout(scope);
  const root = document.createElement('div');
  root.className = 'tx-body';
  const [l, t, r, b] = layout.insets.map(pt);
  root.style.padding = `${t}px ${r}px ${b}px ${l}px`;
  root.style.justifyContent = layout.anchor === 'middle' ? 'center' : layout.anchor === 'bottom' ? 'flex-end' : 'flex-start';
  if (layout.vertical && layout.vertical !== 'horz') root.style.writingMode = layout.vertical === 'vert270' ? 'sideways-lr' : 'vertical-rl';
  const content = document.createElement('div');
  content.className = 'tx-content';
  if (!layout.wrap) content.style.whiteSpace = 'pre';
  root.appendChild(content);

  const paras = children(txBody, 'a:p');
  const empty = !hasText(txBody);
  const usePrompt = empty && opts.prompt && hasText(opts.prompt);
  const source = usePrompt ? children(opts.prompt!, 'a:p') : paras;
  if (usePrompt) root.classList.add('ph-prompt');

  const counters: number[] = [];
  source.forEach((p, pi) => {
    const pPr = child(p, 'a:pPr');
    const lvl = num(pPr, 'lvl', 0);
    const ps = paraStyle(scope, pPr, lvl);
    const el = document.createElement('p');
    el.className = 'tx-p';
    el.dataset.pi = String(pi);
    el.dataset.lvl = String(lvl);
    el.style.textAlign = ({ ctr: 'center', r: 'right', just: 'justify', dist: 'justify' } as Record<string, string>)[ps.algn] ?? 'left';
    el.style.marginLeft = pt(ps.marL) + 'px';
    el.style.textIndent = pt(ps.indent) + 'px';

    const runs = children(p).filter((c) => ['r', 'fld', 'br'].includes(c.localName));
    const firstRPr = child(runs.find((x) => x.localName !== 'br') ?? null, 'a:rPr') ?? child(p, 'a:endParaRPr');
    const base = runStyle(scope, firstRPr, lvl, pPr);
    const fs = base.sz * layout.fontScale;
    el.style.fontSize = fs + 'px';
    el.style.fontFamily = fontStack(base.fonts);
    el.style.color = base.color;
    const lnRed = 1 - layout.lnSpcReduction;
    if (ps.lnSpcPts != null) el.style.lineHeight = ps.lnSpcPts * lnRed + 'px';
    else el.style.lineHeight = String(1.2 * (ps.lnSpcPct ?? 1) * lnRed);
    el.style.marginTop = (pi === 0 ? 0 : ps.spcBefPts + ps.spcBefPct * fs) + 'px';
    el.style.marginBottom = ps.spcAftPts + ps.spcAftPct * fs + 'px';

    for (let k = lvl + 1; k < counters.length; k++) counters[k] = 0;
    const hasContent = runs.some((x) => x.localName !== 'br');
    if (ps.bullet.kind !== 'none' && hasContent) {
      let txt: string;
      if (ps.bullet.kind === 'num') {
        counters[lvl] = (counters[lvl] || ps.bullet.start - 1) + 1;
        txt = autoNum(ps.bullet.scheme, counters[lvl]);
      } else {
        counters[lvl] = 0;
        txt = ps.bullet.char;
      }
      el.classList.add('has-bullet');
      el.style.setProperty('--bu', JSON.stringify(txt));
      el.style.setProperty('--bu-w', Math.max(0, -pt(ps.indent)) + 'px');
      if (ps.bullet.color) el.style.setProperty('--bu-color', ps.bullet.color);
      if (ps.bullet.font) el.style.setProperty('--bu-font', fontStack([ps.bullet.font]));
      el.style.setProperty('--bu-size', ps.bullet.size * 100 + '%');
    } else counters[lvl] = 0;

    runs.forEach((run, ri) => {
      if (run.localName === 'br') {
        const br = document.createElement('br');
        br.dataset.ri = String(ri);
        el.appendChild(br);
        return;
      }
      const text = child(run, 'a:t')?.textContent ?? '';
      const rs = runStyle(scope, child(run, 'a:rPr'), lvl, pPr);
      const span = document.createElement('span');
      span.dataset.ri = String(ri);
      span.textContent = rs.caps ? text.toUpperCase() : text;
      applyRunCss(span, rs, layout.fontScale);
      el.appendChild(span);
    });
    if (!hasContent || runs.at(-1)?.localName === 'br') el.appendChild(document.createElement('br'));
    content.appendChild(el);
  });
  return { root, content, paras, layout, isEmpty: empty };
}

export function applyRunCss(span: HTMLElement, rs: RunStyle, scale: number) {
  span.style.fontSize = rs.sz * scale * (rs.baseline ? 0.7 : 1) + 'px';
  span.style.fontFamily = fontStack(rs.fonts);
  span.style.color = rs.color;
  if (rs.b) span.style.fontWeight = 'bold';
  if (rs.i) span.style.fontStyle = 'italic';
  const deco = [rs.u ? 'underline' : '', rs.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) span.style.textDecorationLine = deco;
  if (rs.baseline > 0) span.style.verticalAlign = 'super';
  if (rs.baseline < 0) span.style.verticalAlign = 'sub';
  if (rs.highlight) span.style.backgroundColor = rs.highlight;
}

// ------------------------------------------------------------ DOM -> XML

function decorations(el: Element, stop: Element) {
  let u = false, s = false;
  for (let e: Element | null = el; e; e = e.parentElement) {
    const line = getComputedStyle(e).textDecorationLine;
    if (line.includes('underline')) u = true;
    if (line.includes('line-through')) s = true;
    if (e.tagName === 'U' || e.tagName === 'INS') u = true;
    if (e.tagName === 'S' || e.tagName === 'STRIKE' || e.tagName === 'DEL') s = true;
    if (e === stop) break;
  }
  return { u, s };
}

function firstFamily(f: string) {
  return f.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
}

/** Compare the computed style of `el` with the reference element and return changed props. */
function diffStyle(el: Element, ref: Element, block: Element, scale: number): RunProps {
  const a = getComputedStyle(el), b = getComputedStyle(ref);
  const out: RunProps = {};
  const bold = (w: string) => Number(w) >= 600 || w === 'bold';
  if (bold(a.fontWeight) !== bold(b.fontWeight)) out.b = bold(a.fontWeight);
  if ((a.fontStyle === 'italic') !== (b.fontStyle === 'italic')) out.i = a.fontStyle === 'italic';
  const da = decorations(el, block), db = decorations(ref, block);
  if (da.u !== db.u) out.u = da.u;
  if (da.s !== db.s) out.strike = da.s;
  const fa = parseFloat(a.fontSize), fb = parseFloat(b.fontSize);
  if (Math.abs(fa - fb) > 0.05) out.sz = Math.round((fa / scale) * 2) / 2;
  if (a.color !== b.color) { const h = cssToHex(a.color); if (h) out.color = h; }
  if (firstFamily(a.fontFamily) !== firstFamily(b.fontFamily)) out.font = firstFamily(a.fontFamily);
  return out;
}

function makeRPr(doc: XMLDocument, from: Element | null): Element {
  if (from) {
    const c = doc.createElementNS(NS.a, 'a:rPr');
    for (const at of Array.from(from.attributes)) c.setAttribute(at.name, at.value);
    for (const ch of Array.from(from.childNodes)) c.appendChild(doc.importNode(ch, true));
    return c;
  }
  return create(doc, 'a:rPr', { lang: 'ko-KR', altLang: 'en-US' });
}

/**
 * Convert edited DOM paragraphs back to a:p elements.
 * Original run properties are kept, and only formatting that changed in the DOM is written explicitly,
 * so inherited styles (theme fonts, placeholder sizes) are preserved.
 */
export function domToParagraphs(content: HTMLElement, origParas: Element[], doc: XMLDocument, scale: number): Element[] {
  const blocks: HTMLElement[] = [];
  let loose: HTMLElement | null = null;
  for (const n of Array.from(content.childNodes)) {
    if (n instanceof HTMLElement && (n.tagName === 'P' || n.tagName === 'DIV')) { blocks.push(n); loose = null; continue; }
    if (n.nodeType === Node.TEXT_NODE && !n.textContent) continue;
    if (!loose) { loose = document.createElement('p'); blocks.push(loose); }
    loose.appendChild(n.cloneNode(true));
  }

  const out: Element[] = [];
  for (const block of blocks) {
    const orig = origParas[Number(block.dataset.pi ?? 0)] ?? origParas[0] ?? null;
    const p = create(doc, 'a:p');
    const origPPr = child(orig, 'a:pPr');
    if (origPPr) p.appendChild(doc.importNode(origPPr, true));
    const algn = block.dataset.algn;
    const lvl = block.dataset.lvlSet;
    if (algn || lvl != null) {
      let pPr = child(p, 'a:pPr');
      if (!pPr) { pPr = create(doc, 'a:pPr'); p.appendChild(pPr); }
      if (algn) pPr.setAttribute('algn', algn);
      if (lvl != null) { if (lvl === '0') pPr.removeAttribute('lvl'); else pPr.setAttribute('lvl', lvl); }
    }
    const origRuns = orig ? children(orig).filter((c) => ['r', 'fld', 'br'].includes(c.localName)) : [];
    const endRPr = child(orig, 'a:endParaRPr');
    let lastRPr: Element | null = null;

    const emit = (text: string | null, holder: Element) => {
      const riEl = holder.closest('[data-ri]');
      const ref = riEl && block.contains(riEl) ? riEl : block;
      const origRun = riEl && block.contains(riEl) ? origRuns[Number((riEl as HTMLElement).dataset.ri)] : null;
      const baseRPr = origRun && origRun.localName !== 'br' ? child(origRun, 'a:rPr') : lastRPr ?? endRPr;
      const rPr = makeRPr(doc, baseRPr);
      setRunProps(rPr, diffStyle(holder, ref, block, scale));
      if (text === null) {
        const br = create(doc, 'a:br');
        br.appendChild(rPr);
        p.appendChild(br);
        return;
      }
      lastRPr = rPr;
      // Merge with the previous run when properties are identical.
      const prev = p.lastElementChild;
      if (prev?.localName === 'r' && new XMLSerializer().serializeToString(child(prev, 'a:rPr')!) === new XMLSerializer().serializeToString(rPr)) {
        const t = child(prev, 'a:t')!;
        t.textContent = (t.textContent ?? '') + text;
        return;
      }
      const r = create(doc, 'a:r');
      r.appendChild(rPr);
      const t = create(doc, 'a:t');
      t.textContent = text;
      r.appendChild(t);
      p.appendChild(r);
    };

    const walk = (node: Node) => {
      for (const n of Array.from(node.childNodes)) {
        if (n.nodeType === Node.TEXT_NODE) {
          const text = (n.textContent ?? '').replace(/​/g, '');
          if (text) emit(text, n.parentElement!);
        } else if (n instanceof HTMLElement) {
          if (n.tagName === 'BR') {
            const isTrailing = !n.nextSibling || (n === block.lastChild);
            if (!isTrailing) emit(null, n.parentElement!);
          } else walk(n);
        }
      }
    };
    walk(block);

    const end = endRPr ? doc.importNode(endRPr, true) as Element : (() => {
      const e = create(doc, 'a:endParaRPr', { lang: 'ko-KR', altLang: 'en-US' });
      if (lastRPr) for (const at of Array.from((lastRPr as Element).attributes)) e.setAttribute(at.name, at.value);
      return e;
    })();
    p.appendChild(end);
    out.push(p);
  }
  if (!out.length) {
    const p = create(doc, 'a:p');
    const e = child(origParas[0] ?? null, 'a:endParaRPr');
    p.appendChild(e ? doc.importNode(e, true) : create(doc, 'a:endParaRPr', { lang: 'ko-KR' }));
    out.push(p);
  }
  return out;
}
