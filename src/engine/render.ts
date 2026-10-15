// Slide renderer: builds a DOM tree (1 CSS px = 1 pt = 12700 EMU) for a slide,
// mirroring HOP's renderPage* API surface but targeting HTML/SVG instead of canvas.
import { PptxDocument, SlideContext, Xfrm, findInherited, placeholder, readXfrm, shapeChildren, shapeId, spTree, xfrmEl } from './document';
import { OPEN_PRESETS, customPath, presetPath } from './geometry';
import { REL } from './package';
import { ColorContext, resolveColor, resolveHex } from './theme';
import { RenderedText, hasText, renderText, textScope } from './text';
import { NS, attr, child, children, num, parseXml, path } from './xml';

const px = (emu: number) => emu / 12700;
let uid = 0;

export class ImageCache {
  private urls = new Map<string, string>();
  constructor(private doc: PptxDocument) {}

  url(part: string): string | null { return this.urls.get(part) ?? null; }

  async ensure(parts: string[]) {
    await Promise.all(parts.filter((p) => !this.urls.has(p)).map(async (p) => {
      const data = await this.doc.pkg.getBinary(p);
      if (!data) return;
      const ext = p.split('.').pop()!.toLowerCase();
      const mime = ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
      this.urls.set(p, URL.createObjectURL(new Blob([data as BlobPart], { type: mime })));
    }));
  }

  /** Load every image referenced by a slide and the layout/master it uses. */
  async ensureForSlide(slidePath: string) {
    const ctx = this.doc.context(slidePath);
    const parts = [slidePath, ctx.layoutPath, ctx.masterPath].filter((p): p is string => !!p);
    const imgs = parts.flatMap((p) => this.doc.pkg.rels(p).filter((r) => r.type === REL.image && !r.external).map((r) => r.target));
    await this.ensure(imgs);
  }

  dispose() {
    for (const u of this.urls.values()) URL.revokeObjectURL(u);
    this.urls.clear();
  }
}

export interface TextTarget {
  txBody: Element;
  rendered: RenderedText;
  cell?: { r: number; c: number };
}

export interface RenderedShape {
  id: string;
  el: HTMLElement;
  shape: Element;
  xfrm: Xfrm;
  kind: string;
  text: TextTarget | null;
  cells: TextTarget[];
  locked: boolean;
  isLine: boolean;
}

export interface RenderedSlide {
  el: HTMLElement;
  shapes: RenderedShape[];
  width: number;
  height: number;
}

export interface RenderOptions {
  editor?: boolean; // show placeholder prompts
}

interface Env {
  doc: PptxDocument;
  ctx: SlideContext;
  images: ImageCache;
  part: string; // part owning the elements being rendered (for r:embed lookup)
  owner: 'slide' | 'layout' | 'master';
  opts: RenderOptions;
  themeDoc: XMLDocument | null;
}

export function renderSlide(doc: PptxDocument, slidePath: string, images: ImageCache, opts: RenderOptions = {}): RenderedSlide {
  const ctx = doc.context(slidePath);
  const { cx, cy } = doc.slideSize;
  const el = document.createElement('div');
  el.className = 'slide';
  el.style.width = px(cx) + 'px';
  el.style.height = px(cy) + 'px';
  const themePath = ctx.masterPath ? doc.pkg.relOfType(ctx.masterPath, REL.theme)?.target : null;
  const themeDoc = themePath ? doc.pkg.getXml(themePath) : null;
  const base: Omit<Env, 'part' | 'owner'> = { doc, ctx, images, opts, themeDoc };

  renderBackground(el, { ...base, part: slidePath, owner: 'slide' });

  // Master and layout decorations (non-placeholder shapes) sit behind slide content.
  const showMaster = (d: XMLDocument | null) => d?.documentElement.getAttribute('showMasterSp') !== '0';
  if (ctx.master && ctx.masterPath && showMaster(ctx.slide) && showMaster(ctx.layout)) {
    renderDecor(el, spTree(ctx.master), { ...base, part: ctx.masterPath, owner: 'master' });
  }
  if (ctx.layout && ctx.layoutPath && showMaster(ctx.slide)) {
    renderDecor(el, spTree(ctx.layout), { ...base, part: ctx.layoutPath, owner: 'layout' });
  }

  const shapes: RenderedShape[] = [];
  const env: Env = { ...base, part: slidePath, owner: 'slide' };
  for (const s of topLevelShapes(spTree(ctx.slide))) {
    const r = renderShape(env, s.el, null, s.locked);
    if (!r) continue;
    r.el.dataset.id = r.id;
    r.el.classList.add('top');
    el.appendChild(r.el);
    shapes.push(r);
  }
  return { el, shapes, width: px(cx), height: px(cy) };
}

function topLevelShapes(tree: Element | null): { el: Element; locked: boolean }[] {
  const out: { el: Element; locked: boolean }[] = [];
  for (const c of children(tree)) {
    if (c.localName === 'AlternateContent') {
      const fb = children(c).find((x) => x.localName === 'Fallback');
      for (const s of children(fb).filter((x) => x.namespaceURI === NS.p)) out.push({ el: s, locked: true });
    } else if (c.namespaceURI === NS.p && ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'].includes(c.localName)) {
      out.push({ el: c, locked: false });
    }
  }
  return out;
}

function renderDecor(slideEl: HTMLElement, tree: Element | null, env: Env) {
  for (const s of topLevelShapes(tree)) {
    if (placeholder(s.el)) continue;
    const r = renderShape(env, s.el, null, true);
    if (r) { r.el.classList.add('decor'); slideEl.appendChild(r.el); }
  }
}

// ------------------------------------------------------------ background

function renderBackground(slideEl: HTMLElement, env: Env) {
  const { ctx } = env;
  const layers: [XMLDocument | null, string | null][] = [[ctx.slide, ctx.slidePath], [ctx.layout, ctx.layoutPath], [ctx.master, ctx.masterPath]];
  for (const [d, part] of layers) {
    const bg = path(d?.documentElement, 'p:cSld', 'p:bg');
    if (!bg || !part) continue;
    const e = { ...env, part };
    const bgPr = child(bg, 'p:bgPr');
    if (bgPr) { slideEl.style.background = fillCss(bgPr, e, env.ctx.color) ?? '#fff'; return; }
    const ref = child(bg, 'p:bgRef');
    if (ref) {
      const idx = num(ref, 'idx');
      const phClr = resolveHex(ref, ctx.color);
      const lst = path(env.themeDoc?.documentElement, 'a:themeElements', 'a:fmtScheme', 'a:bgFillStyleLst');
      const fill = idx >= 1001 ? children(lst)[idx - 1001] : null;
      const colorCtx = { ...ctx.color, phClr };
      if (fill) {
        const holder = fill.ownerDocument.createElementNS(NS.a, 'a:holder');
        holder.appendChild(fill.cloneNode(true));
        slideEl.style.background = fillCss(holder, e, colorCtx) ?? `#${phClr}`;
      } else slideEl.style.background = phClr ? `#${phClr}` : '#fff';
      return;
    }
  }
  slideEl.style.background = '#fff';
}

/** CSS background for a container holding one DrawingML fill element. */
function fillCss(container: Element, env: Env, color: ColorContext): string | null {
  const solid = child(container, 'a:solidFill');
  if (solid) return resolveColor(solid, color);
  const grad = child(container, 'a:gradFill');
  if (grad) {
    const stops = children(path(grad, 'a:gsLst'), 'a:gs').map((g) => `${resolveColor(g, color)} ${num(g, 'pos') / 1000}%`);
    const lin = child(grad, 'a:lin');
    if (child(grad, 'a:path')) return `radial-gradient(${stops.join(',')})`;
    const ang = lin ? num(lin, 'ang') / 60000 + 90 : 180;
    return `linear-gradient(${ang}deg, ${stops.join(',')})`;
  }
  const blip = child(container, 'a:blipFill');
  if (blip) {
    const url = blipUrl(blip, env);
    return url ? `center / cover no-repeat url("${url}")` : null;
  }
  if (child(container, 'a:noFill')) return 'transparent';
  return null;
}

function blipUrl(blipFill: Element, env: Env): string | null {
  const blip = child(blipFill, 'a:blip');
  const id = blip?.getAttributeNS(NS.r, 'embed');
  const rel = env.doc.pkg.rel(env.part, id ?? null);
  return rel ? env.images.url(rel.target) : null;
}

// ------------------------------------------------------------ shapes

interface Paint { fill: string; stroke: string; strokeWidth: number; dash: string | null; gradient: Element | null; image: string | null; headEnd: string | null; tailEnd: string | null }

function shapePaint(env: Env, shape: Element, spPr: Element | null, color: ColorContext): Paint {
  const style = child(shape, 'p:style');
  const p: Paint = { fill: 'none', stroke: 'none', strokeWidth: 0, dash: null, gradient: null, image: null, headEnd: null, tailEnd: null };
  // fill
  const fillRef = child(style, 'a:fillRef');
  const refColor = fillRef ? resolveHex(fillRef, color) : null;
  const fctx = { ...color, phClr: refColor };
  let fillEl: Element | null = null;
  for (const t of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill']) { fillEl = child(spPr, t); if (fillEl) break; }
  if (fillEl) {
    if (fillEl.localName === 'solidFill') p.fill = resolveColor(fillEl, fctx) ?? 'none';
    else if (fillEl.localName === 'gradFill') p.gradient = fillEl;
    else if (fillEl.localName === 'blipFill') p.image = blipUrl(fillEl, env);
    else if (fillEl.localName === 'pattFill') p.fill = resolveColor(child(fillEl, 'a:fgClr'), fctx) ?? 'none';
  } else if (fillRef && num(fillRef, 'idx') > 0 && refColor) p.fill = '#' + refColor;
  // outline
  const ln = child(spPr, 'a:ln');
  const lnRef = child(style, 'a:lnRef');
  const lnRefIdx = num(lnRef, 'idx');
  const lctx = { ...color, phClr: lnRef ? resolveHex(lnRef, color) : null };
  if (ln && child(ln, 'a:noFill')) p.stroke = 'none';
  else if (ln && child(ln, 'a:solidFill')) p.stroke = resolveColor(child(ln, 'a:solidFill'), lctx) ?? 'none';
  else if (ln && child(ln, 'a:gradFill')) p.stroke = resolveColor(path(ln, 'a:gradFill', 'a:gsLst', 'a:gs'), lctx) ?? 'none';
  else if (lnRefIdx > 0 && lctx.phClr) p.stroke = '#' + lctx.phClr;
  const w = ln?.getAttribute('w');
  p.strokeWidth = w != null ? px(Number(w)) : lnRefIdx > 0 ? px(env.ctx.theme.lineWidths[lnRefIdx - 1] ?? 9525) : px(9525);
  if (p.stroke !== 'none' && p.strokeWidth === 0) p.strokeWidth = 0.75;
  const dash = attr(child(ln, 'a:prstDash'), 'val');
  if (dash && dash !== 'solid') {
    const k = Math.max(1, p.strokeWidth);
    p.dash = ({ dot: `${k},${k}`, sysDot: `${k},${k}`, dash: `${4 * k},${3 * k}`, sysDash: `${3 * k},${k}`, lgDash: `${8 * k},${3 * k}`, dashDot: `${4 * k},${3 * k},${k},${3 * k}` } as Record<string, string>)[dash] ?? `${4 * k},${3 * k}`;
  }
  const end = (t: string) => { const v = attr(child(ln, t), 'type'); return v && v !== 'none' ? v : null; };
  p.headEnd = end('a:headEnd');
  p.tailEnd = end('a:tailEnd');
  return p;
}

function svgEl<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number> = {}): SVGElementTagNameMap[K] {
  const e = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
}

function buildSvg(env: Env, shape: Element, spPr: Element | null, w: number, h: number, flipH: boolean, flipV: boolean, color: ColorContext): SVGSVGElement | null {
  const geom = child(spPr, 'a:prstGeom');
  const cust = child(spPr, 'a:custGeom');
  const prst = geom?.getAttribute('prst') ?? (cust ? null : 'rect');
  const paint = shapePaint(env, shape, spPr, color);
  if (paint.fill === 'none' && paint.stroke === 'none' && !paint.gradient && !paint.image) return null;
  const d = cust ? customPath(cust, w, h) : presetPath(prst!, w, h, child(geom, 'a:avLst'));
  const open = !!prst && OPEN_PRESETS.has(prst);
  const svg = svgEl('svg', { width: Math.max(w, 1), height: Math.max(h, 1), class: 'geom' });
  svg.style.overflow = 'visible';
  if (flipH || flipV) svg.style.transform = `scale(${flipH ? -1 : 1},${flipV ? -1 : 1})`;
  const defs = svgEl('defs');
  svg.appendChild(defs);
  const pathEl = svgEl('path', { d });
  let fill = open ? 'none' : paint.fill;
  if (!open && paint.gradient) {
    const id = 'g' + ++uid;
    const lin = child(paint.gradient, 'a:lin');
    const ang = ((lin ? num(lin, 'ang') : 5400000) / 60000) * (Math.PI / 180);
    const isPath = !!child(paint.gradient, 'a:path');
    const g = isPath ? svgEl('radialGradient', { id }) : svgEl('linearGradient', {
      id, x1: 0.5 - Math.cos(ang) / 2, y1: 0.5 - Math.sin(ang) / 2, x2: 0.5 + Math.cos(ang) / 2, y2: 0.5 + Math.sin(ang) / 2,
    });
    for (const gs of children(path(paint.gradient, 'a:gsLst'), 'a:gs')) {
      g.appendChild(svgEl('stop', { offset: num(gs, 'pos') / 100000, 'stop-color': resolveColor(gs, color) ?? '#000' }));
    }
    defs.appendChild(g);
    fill = `url(#${id})`;
  }
  if (!open && paint.image) {
    const id = 'c' + ++uid;
    const clip = svgEl('clipPath', { id });
    clip.appendChild(svgEl('path', { d }));
    defs.appendChild(clip);
    const img = svgEl('image', { href: paint.image, width: w, height: h, preserveAspectRatio: 'none', 'clip-path': `url(#${id})` });
    svg.appendChild(img);
    fill = 'none';
  }
  pathEl.setAttribute('fill', fill);
  pathEl.setAttribute('stroke', paint.stroke);
  pathEl.setAttribute('stroke-width', String(paint.strokeWidth));
  pathEl.setAttribute('stroke-linejoin', 'round');
  if (paint.dash) pathEl.setAttribute('stroke-dasharray', paint.dash);
  if (open && paint.stroke !== 'none') {
    for (const [which, kind] of [['marker-start', paint.headEnd], ['marker-end', paint.tailEnd]] as const) {
      if (!kind) continue;
      const id = 'm' + ++uid;
      const m = svgEl('marker', { id, viewBox: '0 0 10 10', refX: 5, refY: 5, markerWidth: 4, markerHeight: 4, orient: 'auto-start-reverse' });
      m.appendChild(kind === 'oval' ? svgEl('circle', { cx: 5, cy: 5, r: 4, fill: paint.stroke }) : svgEl('path', { d: 'M0,0 L10,5 L0,10 z', fill: paint.stroke }));
      defs.appendChild(m);
      pathEl.setAttribute(which, `url(#${id})`);
    }
  }
  svg.appendChild(pathEl);
  return svg;
}

type Mapper = (x: Xfrm) => Xfrm;

function renderShape(env: Env, shape: Element, map: Mapper | null, locked: boolean): RenderedShape | null {
  const raw = env.owner === 'slide' ? env.doc.effectiveXfrm(env.ctx, shape) : readXfrm(xfrmEl(shape)) ?? inheritedXfrm(env, shape);
  if (!raw) return null;
  const x = map ? map(raw) : raw;
  const el = document.createElement('div');
  el.className = 'shape';
  el.style.left = px(x.x) + 'px';
  el.style.top = px(x.y) + 'px';
  el.style.width = px(x.w) + 'px';
  el.style.height = px(x.h) + 'px';
  if (x.rot) el.style.transform = `rotate(${x.rot}deg)`;
  const w = px(x.w), h = px(x.h);
  const out: RenderedShape = { id: shapeId(shape), el, shape, xfrm: raw, kind: shape.localName, text: null, cells: [], locked, isLine: false };
  const color = env.ctx.color;

  switch (shape.localName) {
    case 'sp': case 'cxnSp': {
      const spPr = child(shape, 'p:spPr');
      const prst = attr(child(spPr, 'a:prstGeom'), 'prst');
      out.isLine = shape.localName === 'cxnSp' || (!!prst && OPEN_PRESETS.has(prst));
      const svg = buildSvg(env, shape, spPr, w, h, x.flipH, x.flipV, color);
      if (svg) el.appendChild(svg);
      const txBody = child(shape, 'p:txBody');
      if (txBody) {
        const scope = textScope(env.ctx, env.doc.pres, shape, txBody, env.owner);
        const ph = placeholder(shape);
        let prompt: Element | null = null;
        if (ph && env.owner === 'slide' && env.opts.editor && !hasText(txBody)) {
          prompt = promptBody(ph, findInherited(ph, env.ctx.layout, false));
        }
        // Layout/master placeholders are never drawn; slide placeholders without text show nothing outside the editor.
        const rendered = renderText(scope, txBody, { prompt });
        if (!rendered.isEmpty || prompt || env.owner === 'slide') {
          el.appendChild(rendered.root);
          out.text = { txBody, rendered };
        }
        if (rendered.layout.autofit === 'shape') el.classList.add('autofit');
      }
      break;
    }
    case 'pic': {
      const blipFill = child(shape, 'p:blipFill');
      const url = blipFill ? blipUrl(blipFill, env) : null;
      const box = document.createElement('div');
      box.className = 'pic';
      if (url) {
        const img = document.createElement('img');
        img.src = url;
        img.draggable = false;
        const src = child(blipFill, 'a:srcRect');
        const l = num(src, 'l') / 100000, t = num(src, 't') / 100000, r = num(src, 'r') / 100000, b = num(src, 'b') / 100000;
        const iw = w / (1 - l - r || 1), ih = h / (1 - t - b || 1);
        img.style.cssText = `position:absolute;left:${-l * iw}px;top:${-t * ih}px;width:${iw}px;height:${ih}px`;
        if (x.flipH || x.flipV) img.style.transform = `scale(${x.flipH ? -1 : 1},${x.flipV ? -1 : 1})`;
        box.appendChild(img);
      } else {
        box.classList.add('missing');
        box.textContent = child(path(shape, 'p:nvPicPr', 'p:nvPr'), 'a:videoFile') ? '동영상' : '그림';
      }
      const geom = attr(path(shape, 'p:spPr', 'a:prstGeom'), 'prst');
      if (geom === 'ellipse') box.style.borderRadius = '50%';
      el.appendChild(box);
      const ln = path(shape, 'p:spPr', 'a:ln');
      if (ln && !child(ln, 'a:noFill')) {
        const p = shapePaint(env, shape, child(shape, 'p:spPr'), color);
        if (p.stroke !== 'none') box.style.outline = `${p.strokeWidth}px solid ${p.stroke}`;
      }
      break;
    }
    case 'grpSp': {
      const gx = readXfrm(xfrmEl(shape));
      const xf = xfrmEl(shape);
      const chOff = child(xf, 'a:chOff'), chExt = child(xf, 'a:chExt');
      const cox = num(chOff, 'x'), coy = num(chOff, 'y');
      const sx = gx && num(chExt, 'cx') ? gx.w / num(chExt, 'cx') : 1;
      const sy = gx && num(chExt, 'cy') ? gx.h / num(chExt, 'cy') : 1;
      // Children are placed relative to the group box (whose own position is already applied).
      const childMap: Mapper = (c) => ({ ...c, x: (c.x - cox) * sx, y: (c.y - coy) * sy, w: c.w * sx, h: c.h * sy });
      for (const c of shapeChildren(shape)) {
        const r = renderShape(env, c, childMap, true);
        if (r) el.appendChild(r.el);
      }
      el.classList.add('group');
      break;
    }
    case 'graphicFrame': {
      const data = path(shape, 'a:graphic', 'a:graphicData');
      const tbl = child(data, 'a:tbl');
      if (tbl) renderTable(env, shape, tbl, el, out);
      else {
        const uri = data?.getAttribute('uri') ?? '';
        const box = document.createElement('div');
        box.className = 'frame-ph';
        box.textContent = uri.includes('chart') ? '📊 차트' : uri.includes('diagram') ? 'SmartArt' : '개체';
        el.appendChild(box);
      }
      break;
    }
  }
  return out;
}

const PROMPTS: Record<string, string> = {
  title: '제목을 입력하려면 클릭하십시오', ctrTitle: '제목을 입력하려면 클릭하십시오',
  subTitle: '부제목을 입력하려면 클릭하십시오', pic: '그림을 넣으려면 아이콘을 클릭하십시오',
};

/** Prompt text for an empty placeholder: the layout's custom prompt, else a generic one. */
function promptBody(ph: Element, layoutPh: Element | null): Element | null {
  const layoutTx = child(layoutPh, 'p:txBody');
  if (layoutTx && placeholder(layoutPh!)?.getAttribute('hasCustomPrompt') === '1') return layoutTx;
  const type = ph.getAttribute('type') || 'obj';
  if (['dt', 'ftr', 'sldNum', 'hdr'].includes(type)) return null;
  const text = PROMPTS[type] ?? '텍스트를 입력하려면 클릭하십시오';
  return parseXml(`<p:txBody xmlns:a="${NS.a}" xmlns:p="${NS.p}"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR"/><a:t>${text}</a:t></a:r></a:p></p:txBody>`).documentElement;
}

function inheritedXfrm(env: Env, shape: Element): Xfrm | null {
  const ph = placeholder(shape);
  if (!ph || env.owner !== 'layout') return null;
  const m = findInherited(ph, env.ctx.master, true);
  return m ? readXfrm(xfrmEl(m)) : null;
}

// ------------------------------------------------------------ tables

interface CellStyle { fill: string | null; color: string | null; bold: boolean; border: string | null }

const DEFAULT_TABLE_STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

function tableStyle(env: Env, tbl: Element) {
  const tblPr = child(tbl, 'a:tblPr');
  const id = child(tblPr, 'a:tableStyleId')?.textContent?.trim() ?? null;
  const flags = {
    firstRow: tblPr?.getAttribute('firstRow') === '1', lastRow: tblPr?.getAttribute('lastRow') === '1',
    firstCol: tblPr?.getAttribute('firstCol') === '1', lastCol: tblPr?.getAttribute('lastCol') === '1',
    bandRow: tblPr?.getAttribute('bandRow') === '1', bandCol: tblPr?.getAttribute('bandCol') === '1',
  };
  const stylesPath = env.doc.pkg.rels(env.doc.presPath).find((r) => r.type.endsWith('/tableStyles'))?.target;
  const styles = stylesPath ? env.doc.pkg.getXml(stylesPath) : null;
  const def = children(styles?.documentElement, 'a:tblStyle').find((s) => s.getAttribute('styleId') === id) ?? null;
  const color = env.ctx.color;
  const part = (name: string): Partial<CellStyle> | null => {
    if (def) {
      const e = child(def, `a:${name}`);
      if (!e) return null;
      const fill = path(e, 'a:tcStyle', 'a:fill');
      const tx = child(e, 'a:tcTxStyle');
      const borders = path(e, 'a:tcStyle', 'a:tcBdr');
      const anyLn = borders ? children(borders).map((b) => child(b, 'a:ln')).find((l) => l && child(l, 'a:solidFill')) : null;
      return {
        fill: fill ? (child(fill, 'a:noFill') ? 'transparent' : resolveColor(child(fill, 'a:solidFill'), color)) : undefined as unknown as null,
        color: tx ? resolveColor(tx, color) : undefined as unknown as null,
        bold: tx?.getAttribute('b') === 'on' ? true : undefined as unknown as boolean,
        border: anyLn ? `${px(num(anyLn, 'w', 12700))}px solid ${resolveColor(child(anyLn, 'a:solidFill'), color)}` : undefined as unknown as null,
      };
    }
    if (id !== DEFAULT_TABLE_STYLE && id != null) return null;
    const a1 = color.theme.colors.accent1;
    const tint = (k: number) => {
      const n = parseInt(a1, 16);
      const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * (1 - k)));
      return `rgb(${c.join(',')})`;
    };
    switch (name) {
      case 'wholeTbl': return { fill: tint(0.2), color: '#000', border: '1px solid #fff' };
      case 'band1H': return { fill: tint(0.4) };
      case 'firstRow': case 'lastRow': return { fill: '#' + a1, color: '#fff', bold: true };
      case 'firstCol': case 'lastCol': return { bold: true };
      default: return null;
    }
  };
  return (r: number, c: number, rows: number, cols: number): CellStyle => {
    const s: CellStyle = { fill: null, color: null, bold: false, border: null };
    const merge = (p: Partial<CellStyle> | null) => {
      if (!p) return;
      for (const k of Object.keys(p) as (keyof CellStyle)[]) if (p[k] !== undefined) (s as unknown as Record<string, unknown>)[k] = p[k];
    };
    merge(part('wholeTbl'));
    const bodyRow = flags.firstRow ? r - 1 : r;
    if (flags.bandRow && bodyRow >= 0 && bodyRow % 2 === 0 && !(flags.firstRow && r === 0)) merge(part('band1H'));
    if (flags.bandCol && c % 2 === 0) merge(part('band1V'));
    if (flags.firstCol && c === 0) merge(part('firstCol'));
    if (flags.lastCol && c === cols - 1) merge(part('lastCol'));
    if (flags.lastRow && r === rows - 1) merge(part('lastRow'));
    if (flags.firstRow && r === 0) merge(part('firstRow'));
    return s;
  };
}

function renderTable(env: Env, shape: Element, tbl: Element, host: HTMLElement, out: RenderedShape) {
  const table = document.createElement('table');
  table.className = 'tbl';
  const cols = children(child(tbl, 'a:tblGrid'), 'a:gridCol').map((g) => px(num(g, 'w')));
  const colgroup = document.createElement('colgroup');
  for (const w of cols) { const c = document.createElement('col'); c.style.width = w + 'px'; colgroup.appendChild(c); }
  table.appendChild(colgroup);
  const rows = children(tbl, 'a:tr');
  const style = tableStyle(env, tbl);
  rows.forEach((tr, r) => {
    const rowEl = document.createElement('tr');
    rowEl.style.height = px(num(tr, 'h')) + 'px';
    children(tr, 'a:tc').forEach((tc, c) => {
      if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') return;
      const td = document.createElement('td');
      const gs = num(tc, 'gridSpan', 1), rs = num(tc, 'rowSpan', 1);
      if (gs > 1) td.colSpan = gs;
      if (rs > 1) td.rowSpan = rs;
      const cs = style(r, c, rows.length, cols.length);
      const tcPr = child(tc, 'a:tcPr');
      const ownFill = child(tcPr, 'a:solidFill');
      td.style.background = ownFill ? resolveColor(ownFill, env.ctx.color) ?? '' : child(tcPr, 'a:noFill') ? 'transparent' : cs.fill ?? 'transparent';
      if (cs.border) td.style.border = cs.border;
      for (const [tag, side] of [['a:lnL', 'Left'], ['a:lnR', 'Right'], ['a:lnT', 'Top'], ['a:lnB', 'Bottom']] as const) {
        const ln = child(tcPr, tag);
        if (!ln) continue;
        const k = `border${side}` as 'borderLeft';
        td.style[k] = child(ln, 'a:noFill') ? 'none' : `${px(num(ln, 'w', 12700))}px solid ${resolveColor(child(ln, 'a:solidFill'), env.ctx.color) ?? '#000'}`;
      }
      const anchor = tcPr?.getAttribute('anchor');
      td.style.verticalAlign = anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';
      const txBody = child(tc, 'a:txBody');
      if (txBody) {
        const scope = textScope(env.ctx, env.doc.pres, shape, txBody, env.owner, true);
        if (cs.color) scope.fontRefColor = cs.color;
        const rendered = renderText(scope, txBody);
        rendered.root.style.padding = `${px(num(tcPr, 'marT', 45720))}px ${px(num(tcPr, 'marR', 91440))}px ${px(num(tcPr, 'marB', 45720))}px ${px(num(tcPr, 'marL', 91440))}px`;
        if (cs.bold) rendered.content.style.fontWeight = 'bold';
        td.appendChild(rendered.root);
        td.dataset.cell = `${r},${c}`;
        out.cells.push({ txBody, rendered, cell: { r, c } });
      }
      rowEl.appendChild(td);
    });
    table.appendChild(rowEl);
  });
  host.appendChild(table);
}
