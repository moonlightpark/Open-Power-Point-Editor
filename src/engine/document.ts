// PptxDocument: the presentation model and every mutating command.
// The XML parts are the source of truth (like HOP's rhwp core keeps the HWP model);
// renderers read them and commands mutate them, so untouched content round-trips verbatim.
import { CT, OpcPackage, REL, dirname, relativePath, relsPathFor } from './package';
import { ColorContext, Theme, parseClrMap, parseTheme } from './theme';
import { NS, attr, child, children, create, descendants, ensure, num, parseXml, path, serializeXml } from './xml';

export const EMU_PER_PT = 12700;
export const EMU_PER_CM = 360000;

export interface SlideRef {
  index: number;
  id: string; // p:sldId/@id
  rId: string;
  path: string;
}

export interface Xfrm { x: number; y: number; w: number; h: number; rot: number; flipH: boolean; flipV: boolean }

export interface RunProps {
  b?: boolean; i?: boolean; u?: boolean; strike?: boolean;
  sz?: number; // points
  color?: string; // RRGGBB
  font?: string;
}

export interface SlideContext {
  slidePath: string;
  slide: XMLDocument;
  layoutPath: string | null;
  layout: XMLDocument | null;
  masterPath: string | null;
  master: XMLDocument | null;
  theme: Theme;
  color: ColorContext;
}

interface HistoryEntry { label: string; before: Map<string, string | null>; after: Map<string, string | null> }

const SHAPE_TAGS = new Set(['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp']);

const XMLNS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

export function isShapeEl(el: Element) {
  return el.namespaceURI === NS.p && SHAPE_TAGS.has(el.localName);
}

export function spTree(doc: XMLDocument | null): Element | null {
  return path(doc?.documentElement, 'p:cSld', 'p:spTree');
}

export function shapeChildren(tree: Element | null): Element[] {
  return children(tree).filter(isShapeEl);
}

export function nvPr(shape: Element): Element | null {
  const nv = children(shape).find((c) => c.localName.startsWith('nv'));
  return child(nv, 'p:nvPr');
}

export function cNvPr(shape: Element): Element | null {
  const nv = children(shape).find((c) => c.localName.startsWith('nv'));
  return child(nv, 'p:cNvPr');
}

export function shapeId(shape: Element): string {
  return cNvPr(shape)?.getAttribute('id') ?? '';
}

export function placeholder(shape: Element): Element | null {
  return child(nvPr(shape), 'p:ph');
}

export function xfrmEl(shape: Element, make = false): Element | null {
  switch (shape.localName) {
    case 'graphicFrame': return make ? ensure(shape, 'p:xfrm', ['a:graphic']) : child(shape, 'p:xfrm');
    case 'grpSp': {
      const pr = child(shape, 'p:grpSpPr');
      return pr && make ? ensure(pr, 'a:xfrm') : child(pr, 'a:xfrm');
    }
    default: {
      const pr = child(shape, 'p:spPr');
      if (!pr) return null;
      if (!make) return child(pr, 'a:xfrm');
      const x = child(pr, 'a:xfrm');
      if (x) return x;
      const made = create(shape.ownerDocument, 'a:xfrm');
      pr.insertBefore(made, pr.firstChild);
      return made;
    }
  }
}

export function readXfrm(x: Element | null): Xfrm | null {
  const off = child(x, 'a:off'), ext = child(x, 'a:ext');
  if (!x || !off || !ext) return null;
  return {
    x: num(off, 'x'), y: num(off, 'y'), w: num(ext, 'cx'), h: num(ext, 'cy'),
    rot: num(x, 'rot') / 60000,
    flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1',
  };
}

function phType(ph: Element) {
  return ph.getAttribute('type') || 'obj';
}

const MASTER_TYPE: Record<string, string> = { ctrTitle: 'title', subTitle: 'body', obj: 'body', tbl: 'body', chart: 'body', pic: 'body', dgm: 'body', media: 'body', clipArt: 'body' };

/** Find the placeholder shape a slide/layout placeholder inherits from in `parent`. */
export function findInherited(ph: Element, parent: XMLDocument | null, isMaster: boolean): Element | null {
  const shapes = shapeChildren(spTree(parent)).filter((s) => placeholder(s));
  const type = phType(ph);
  const idx = ph.getAttribute('idx');
  if (!isMaster) {
    if (idx != null) {
      const byIdx = shapes.find((s) => placeholder(s)!.getAttribute('idx') === idx);
      if (byIdx) return byIdx;
    }
    const byType = shapes.find((s) => phType(placeholder(s)!) === type);
    if (byType) return byType;
    if (type === 'ctrTitle' || type === 'title') return shapes.find((s) => ['title', 'ctrTitle'].includes(phType(placeholder(s)!))) ?? null;
    return null;
  }
  const mt = MASTER_TYPE[type] ?? type;
  return shapes.find((s) => phType(placeholder(s)!) === mt) ?? null;
}

export class PptxDocument {
  readonly presPath: string;
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  dirty = false;
  private depth = 0;
  onChange: (() => void) | null = null;

  private constructor(readonly pkg: OpcPackage) {
    const rel = pkg.relOfType('', REL.officeDocument);
    this.presPath = rel?.target ?? 'ppt/presentation.xml';
  }

  static async load(data: ArrayBuffer | Uint8Array) {
    const pkg = await OpcPackage.load(data);
    const d = new PptxDocument(pkg);
    if (!d.pres) throw new Error('presentation.xml을 찾을 수 없습니다. 올바른 PPTX 파일이 아닙니다.');
    return d;
  }

  get pres(): XMLDocument {
    return this.pkg.getXml(this.presPath)!;
  }

  get slideSize() {
    const s = path(this.pres.documentElement, 'p:sldSz');
    return { cx: num(s, 'cx', 12192000), cy: num(s, 'cy', 6858000) };
  }

  slides(): SlideRef[] {
    const lst = path(this.pres.documentElement, 'p:sldIdLst');
    return children(lst, 'p:sldId').map((s, index) => {
      const rId = s.getAttributeNS(NS.r, 'id') || '';
      return { index, id: s.getAttribute('id') || '', rId, path: this.pkg.rel(this.presPath, rId)?.target || '' };
    }).filter((s) => !!s.path && !!this.pkg.getXml(s.path));
  }

  context(slidePath: string): SlideContext {
    const slide = this.pkg.getXml(slidePath)!;
    const layoutPath = this.pkg.relOfType(slidePath, REL.slideLayout)?.target ?? null;
    const layout = layoutPath ? this.pkg.getXml(layoutPath) : null;
    const masterPath = layoutPath ? this.pkg.relOfType(layoutPath, REL.slideMaster)?.target ?? null : null;
    const master = masterPath ? this.pkg.getXml(masterPath) : null;
    const themePath = masterPath ? this.pkg.relOfType(masterPath, REL.theme)?.target : null;
    const theme = parseTheme(themePath ? this.pkg.getXml(themePath) : null);
    let clrMap = parseClrMap(child(master?.documentElement, 'p:clrMap'));
    const ovr = path(slide.documentElement, 'p:clrMapOvr', 'a:overrideClrMapping');
    if (ovr) clrMap = parseClrMap(ovr);
    return { slidePath, slide, layoutPath, layout, masterPath, master, theme, color: { theme, clrMap } };
  }

  layouts(masterPath?: string): { path: string; name: string }[] {
    const masters = masterPath ? [masterPath] : this.pkg.rels(this.presPath).filter((r) => r.type === REL.slideMaster).map((r) => r.target);
    const out: { path: string; name: string }[] = [];
    for (const m of masters) {
      const lst = path(this.pkg.getXml(m)?.documentElement, 'p:sldLayoutIdLst');
      for (const l of children(lst, 'p:sldLayoutId')) {
        const t = this.pkg.rel(m, l.getAttributeNS(NS.r, 'id'))?.target;
        if (!t) continue;
        const name = attr(path(this.pkg.getXml(t)?.documentElement, 'p:cSld'), 'name') || t.split('/').pop()!;
        out.push({ path: t, name });
      }
    }
    return out;
  }

  layoutName(slidePath: string) {
    const lp = this.pkg.relOfType(slidePath, REL.slideLayout)?.target;
    return lp ? attr(path(this.pkg.getXml(lp)?.documentElement, 'p:cSld'), 'name') || '' : '';
  }

  /** Effective xfrm of a shape, following placeholder inheritance into layout/master. */
  effectiveXfrm(ctx: SlideContext, shape: Element, owner: 'slide' | 'layout' | 'master' = 'slide'): Xfrm | null {
    const own = readXfrm(xfrmEl(shape));
    if (own) return own;
    const ph = placeholder(shape);
    if (!ph) return null;
    if (owner === 'slide') {
      const lay = findInherited(ph, ctx.layout, false);
      if (lay) return this.effectiveXfrm(ctx, lay, 'layout');
      const mas = findInherited(ph, ctx.master, true);
      return mas ? readXfrm(xfrmEl(mas)) : null;
    }
    if (owner === 'layout') {
      const mas = findInherited(ph, ctx.master, true);
      return mas ? readXfrm(xfrmEl(mas)) : null;
    }
    return null;
  }

  findShape(slidePath: string, id: string): Element | null {
    const doc = this.pkg.getXml(slidePath);
    if (!doc) return null;
    for (const t of ['p:sp', 'p:pic', 'p:grpSp', 'p:graphicFrame', 'p:cxnSp']) {
      for (const s of descendants(doc, t)) if (shapeId(s) === id) return s;
    }
    return null;
  }

  // ---------------------------------------------------------------- history

  private structuralParts(): string[] {
    const parts = [this.presPath, relsPathFor(this.presPath), '[Content_Types].xml'];
    for (const s of this.slides()) parts.push(s.path, relsPathFor(s.path));
    return parts;
  }

  /**
   * Run a mutation as one undoable step. `parts` lists the XML parts it may touch;
   * 'structure' means the slide list itself (presentation + every slide).
   */
  transact<T>(label: string, parts: string[] | 'structure', fn: () => T): T {
    // Nested transactions fold into the outermost one.
    if (this.depth > 0) return fn();
    this.depth++;
    try {
      return this.transactOuter(label, parts, fn);
    } finally {
      this.depth--;
    }
  }

  private transactOuter<T>(label: string, parts: string[] | 'structure', fn: () => T): T {
    const list = parts === 'structure' ? this.structuralParts() : parts;
    const before = this.pkg.snapshot(list);
    let result: T;
    try {
      result = fn();
    } catch (e) {
      this.pkg.restore(before);
      throw e;
    }
    const afterList = new Set(list);
    if (parts === 'structure') for (const p of this.structuralParts()) afterList.add(p);
    for (const p of afterList) if (!before.has(p)) before.set(p, null);
    const after = this.pkg.snapshot([...afterList]);
    const changed = [...after].some(([p, s]) => before.get(p) !== s);
    if (changed) {
      this.undoStack.push({ label, before, after });
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack = [];
      this.dirty = true;
      this.depth--;
      try { this.onChange?.(); } finally { this.depth++; }
    }
    return result;
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }
  undoLabel() { return this.undoStack.at(-1)?.label; }

  undo() {
    const e = this.undoStack.pop();
    if (!e) return false;
    this.pkg.restore(e.before);
    this.redoStack.push(e);
    this.dirty = true;
    this.onChange?.();
    return true;
  }

  redo() {
    const e = this.redoStack.pop();
    if (!e) return false;
    this.pkg.restore(e.after);
    this.undoStack.push(e);
    this.dirty = true;
    this.onChange?.();
    return true;
  }

  async save(): Promise<Blob> {
    const blob = await this.pkg.save();
    this.dirty = false;
    return blob;
  }

  async saveBytes(): Promise<Uint8Array> {
    const bytes = await this.pkg.saveBytes();
    this.dirty = false;
    return bytes;
  }

  // ---------------------------------------------------------------- slides

  private nextSlideId() {
    const ids = children(path(this.pres.documentElement, 'p:sldIdLst'), 'p:sldId').map((s) => num(s, 'id'));
    return Math.max(255, ...ids) + 1;
  }

  private insertSlidePart(slideDoc: XMLDocument, rels: { type: string; target: string }[], at: number): string {
    const p = this.pkg.uniqueName((n) => `ppt/slides/slide${n}.xml`);
    this.pkg.setXml(p, slideDoc);
    for (const r of rels) this.pkg.addRel(p, r.type, r.target);
    this.pkg.addOverride(p, CT.slide);
    const rId = this.pkg.addRel(this.presPath, REL.slide, p);
    const root = this.pres.documentElement;
    const lst = ensure(root, 'p:sldIdLst', ['p:sldSz', 'p:notesSz']);
    const el = create(this.pres, 'p:sldId', { id: this.nextSlideId() });
    el.setAttributeNS(NS.r, 'r:id', rId);
    const items = children(lst, 'p:sldId');
    lst.insertBefore(el, items[at] ?? null);
    return p;
  }

  addSlide(layoutPath: string, at: number): number {
    return this.transact('새 슬라이드', 'structure', () => {
      const layout = this.pkg.getXml(layoutPath);
      const doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${XMLNS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
      const tree = spTree(doc)!;
      let id = 2;
      for (const s of shapeChildren(spTree(layout))) {
        const ph = placeholder(s);
        if (!ph || ['dt', 'ftr', 'sldNum', 'hdr'].includes(ph.getAttribute('type') || '')) continue;
        const name = cNvPr(s)?.getAttribute('name') || 'Placeholder';
        const sp = parseXml(`<p:sp ${XMLNS}><p:nvSpPr><p:cNvPr id="${id++}" name=""/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR" altLang="en-US"/></a:p></p:txBody></p:sp>`).documentElement;
        cNvPr(sp)!.setAttribute('name', name);
        nvPr(sp)!.appendChild(doc.importNode(ph, true));
        tree.appendChild(doc.importNode(sp, true));
      }
      this.insertSlidePart(doc, [{ type: REL.slideLayout, target: layoutPath }], at);
      return at;
    });
  }

  /** Serialize slides (XML + relationships) so they can be pasted later, even after deletion. */
  copySlides(indices: number[]): ClipSlide[] {
    const all = this.slides();
    return [...indices].sort((a, b) => a - b).map((i) => all[i]).filter(Boolean).map((s) => ({
      xml: serializeXml(this.pkg.getXml(s.path)!),
      rels: children(this.pkg.getXml(relsPathFor(s.path))?.documentElement).map((r) => {
        const external = r.getAttribute('TargetMode') === 'External';
        const rel = this.pkg.rel(s.path, r.getAttribute('Id'));
        return { id: r.getAttribute('Id') || '', type: r.getAttribute('Type') || '', target: external ? r.getAttribute('Target') || '' : rel?.target || '', external };
      }),
    }));
  }

  /** Insert copied slides at position `at`. Returns the index of the first inserted slide. */
  pasteSlides(clip: ClipSlide[], at: number, label = '슬라이드 붙여넣기'): number {
    return this.transact(label, 'structure', () => {
      at = Math.max(0, Math.min(at, this.slides().length));
      clip.forEach((c, k) => {
        const p = this.insertSlidePart(parseXml(c.xml), [], at + k);
        const relDoc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"/>`);
        // Keep the original relationship ids, since the slide XML refers to them.
        // Notes belong to exactly one slide, and parts that no longer exist are dropped.
        for (const r of c.rels) {
          if (r.type === REL.notesSlide || (!r.external && !this.pkg.has(r.target))) continue;
          const el = relDoc.createElementNS(NS.rel, 'Relationship');
          el.setAttribute('Id', r.id);
          el.setAttribute('Type', r.type);
          el.setAttribute('Target', r.external ? r.target : relativePath(dirname(p), r.target));
          if (r.external) el.setAttribute('TargetMode', 'External');
          relDoc.documentElement.appendChild(el);
        }
        this.pkg.setXml(relsPathFor(p), relDoc);
      });
      return at;
    });
  }

  /** Duplicate slides; the copies are placed after the last selected slide. */
  duplicateSlides(indices: number[]): number {
    const sorted = [...indices].sort((a, b) => a - b);
    return this.pasteSlides(this.copySlides(sorted), sorted[sorted.length - 1] + 1, '슬라이드 복제');
  }

  duplicateSlide(index: number): number {
    return this.duplicateSlides([index]);
  }

  deleteSlides(indices: number[]) {
    this.transact('슬라이드 삭제', 'structure', () => {
      const all = this.slides();
      const lst = path(this.pres.documentElement, 'p:sldIdLst')!;
      for (const i of [...indices].sort((a, b) => b - a)) {
        const s = all[i];
        if (!s) continue;
        for (const el of children(lst, 'p:sldId')) if (el.getAttribute('id') === s.id) el.remove();
        this.pkg.removeRel(this.presPath, s.rId);
        this.pkg.removeOverride(s.path);
        this.pkg.remove(s.path);
        this.pkg.remove(relsPathFor(s.path));
      }
    });
  }

  moveSlide(from: number, to: number) {
    if (from === to) return;
    this.transact('슬라이드 이동', [this.presPath], () => {
      const lst = path(this.pres.documentElement, 'p:sldIdLst')!;
      const items = children(lst, 'p:sldId');
      const el = items[from];
      el.remove();
      const rest = children(lst, 'p:sldId');
      lst.insertBefore(el, rest[to] ?? null);
    });
  }

  setSlideLayout(slidePath: string, layoutPath: string) {
    this.transact('레이아웃 변경', [relsPathFor(slidePath)], () => {
      const r = this.pkg.relOfType(slidePath, REL.slideLayout);
      if (r) this.pkg.removeRel(slidePath, r.id);
      this.pkg.addRel(slidePath, REL.slideLayout, layoutPath);
    });
  }

  setSlideBackground(slidePath: string, hex: string | null) {
    this.transact('배경 변경', [slidePath], () => {
      const cSld = path(this.pkg.getXml(slidePath)!.documentElement, 'p:cSld')!;
      child(cSld, 'p:bg')?.remove();
      if (!hex) return;
      const bg = parseXml(`<p:bg ${XMLNS}><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`).documentElement;
      cSld.insertBefore(cSld.ownerDocument.importNode(bg, true), cSld.firstChild);
    });
  }

  notesText(slidePath: string): string {
    const n = this.pkg.relOfType(slidePath, REL.notesSlide)?.target;
    const doc = n ? this.pkg.getXml(n) : null;
    if (!doc) return '';
    for (const sp of descendants(doc, 'p:sp')) {
      if (placeholder(sp)?.getAttribute('type') === 'body') {
        return children(child(sp, 'p:txBody'), 'a:p').map((p) => descendants(p, 'a:t').map((t) => t.textContent).join('')).join('\n');
      }
    }
    return '';
  }

  // ---------------------------------------------------------------- shapes

  private nextShapeId(doc: XMLDocument) {
    let max = 1;
    for (const c of descendants(doc, 'p:cNvPr')) max = Math.max(max, num(c, 'id'));
    return max + 1;
  }

  private appendShape(slidePath: string, xml: string): string {
    const doc = this.pkg.getXml(slidePath)!;
    const el = doc.importNode(parseXml(xml).documentElement, true) as Element;
    const id = String(this.nextShapeId(doc));
    cNvPr(el)!.setAttribute('id', id);
    const tree = spTree(doc)!;
    const ext = child(tree, 'p:extLst');
    tree.insertBefore(el, ext);
    return id;
  }

  addShape(slidePath: string, prst: string, x: Xfrm, label: string): string {
    return this.transact('도형 넣기', [slidePath], () => {
      const xf = `<a:xfrm${x.flipH ? ' flipH="1"' : ''}${x.flipV ? ' flipV="1"' : ''}><a:off x="${Math.round(x.x)}" y="${Math.round(x.y)}"/><a:ext cx="${Math.round(x.w)}" cy="${Math.round(x.h)}"/></a:xfrm>`;
      if (prst === 'line') {
        return this.appendShape(slidePath, `<p:cxnSp ${XMLNS}><p:nvCxnSpPr><p:cNvPr id="0" name="${label}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xf}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>`);
      }
      return this.appendShape(slidePath, `<p:sp ${XMLNS}><p:nvSpPr><p:cNvPr id="0" name="${label}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xf}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="ko-KR" altLang="en-US"/></a:p></p:txBody></p:sp>`);
    });
  }

  addTextBox(slidePath: string, x: Xfrm, text = ''): string {
    return this.transact('글상자 넣기', [slidePath], () => {
      const run = text ? `<a:r><a:rPr lang="ko-KR" altLang="en-US" dirty="0"/><a:t>${escapeXml(text)}</a:t></a:r>` : '';
      return this.appendShape(slidePath, `<p:sp ${XMLNS}><p:nvSpPr><p:cNvPr id="0" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x.x)}" y="${Math.round(x.y)}"/><a:ext cx="${Math.round(x.w)}" cy="${Math.round(x.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>${run}<a:endParaRPr lang="ko-KR" altLang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`);
    });
  }

  addPicture(slidePath: string, data: Uint8Array, mime: string, x: Xfrm, name = '그림'): string {
    const ext = ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/bmp': 'bmp', 'image/webp': 'webp' } as Record<string, string>)[mime] ?? 'png';
    return this.transact('그림 넣기', [slidePath, relsPathFor(slidePath), '[Content_Types].xml'], () => {
      const media = this.pkg.uniqueName((n) => `ppt/media/image${n}.${ext}`);
      this.pkg.setBinary(media, data);
      this.pkg.ensureDefault(ext, mime);
      const rId = this.pkg.addRel(slidePath, REL.image, media);
      return this.appendShape(slidePath, `<p:pic ${XMLNS}><p:nvPicPr><p:cNvPr id="0" name="${escapeXml(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(x.x)}" y="${Math.round(x.y)}"/><a:ext cx="${Math.round(x.w)}" cy="${Math.round(x.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
    });
  }

  addTable(slidePath: string, rows: number, cols: number, x: Xfrm): string {
    return this.transact('표 넣기', [slidePath], () => {
      const cw = Math.round(x.w / cols), rh = 370840;
      const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR" altLang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>';
      const grid = Array.from({ length: cols }, () => `<a:gridCol w="${cw}"/>`).join('');
      const trs = Array.from({ length: rows }, () => `<a:tr h="${rh}">${cell.repeat(cols)}</a:tr>`).join('');
      return this.appendShape(slidePath, `<p:graphicFrame ${XMLNS}><p:nvGraphicFramePr><p:cNvPr id="0" name="표"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round(x.x)}" y="${Math.round(x.y)}"/><a:ext cx="${cw * cols}" cy="${rh * rows}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
    });
  }

  /** Set position/size of shapes. Placeholders get an explicit xfrm (detaching from the layout). */
  setXfrms(slidePath: string, updates: { id: string; x: Xfrm }[], label = '개체 이동') {
    this.transact(label, [slidePath], () => {
      for (const { id, x } of updates) {
        const s = this.findShape(slidePath, id);
        if (!s) continue;
        const xf = xfrmEl(s, true)!;
        const off = ensure(xf, 'a:off');
        const ext = ensure(xf, 'a:ext');
        xf.insertBefore(off, xf.firstChild);
        xf.insertBefore(ext, off.nextSibling);
        off.setAttribute('x', String(Math.round(x.x)));
        off.setAttribute('y', String(Math.round(x.y)));
        ext.setAttribute('cx', String(Math.max(0, Math.round(x.w))));
        ext.setAttribute('cy', String(Math.max(0, Math.round(x.h))));
        if (x.rot) xf.setAttribute('rot', String(Math.round(x.rot * 60000))); else xf.removeAttribute('rot');
        if (x.flipH) xf.setAttribute('flipH', '1'); else xf.removeAttribute('flipH');
        if (x.flipV) xf.setAttribute('flipV', '1'); else xf.removeAttribute('flipV');
        // A group keeps its child coordinate space: chOff/chExt must exist.
        if (s.localName === 'grpSp' && !child(xf, 'a:chOff')) {
          xf.appendChild(create(s.ownerDocument, 'a:chOff', { x: Math.round(x.x), y: Math.round(x.y) }));
          xf.appendChild(create(s.ownerDocument, 'a:chExt', { cx: Math.round(x.w), cy: Math.round(x.h) }));
        }
      }
    });
  }

  deleteShapes(slidePath: string, ids: string[]) {
    this.transact('개체 지우기', [slidePath], () => {
      for (const id of ids) this.findShape(slidePath, id)?.remove();
    });
  }

  /** z-order: 'front' | 'back' | 'forward' | 'backward' */
  reorder(slidePath: string, ids: string[], how: 'front' | 'back' | 'forward' | 'backward') {
    this.transact('순서 바꾸기', [slidePath], () => {
      const tree = spTree(this.pkg.getXml(slidePath))!;
      for (const id of ids) {
        const s = this.findShape(slidePath, id);
        if (!s || s.parentElement !== tree) continue;
        const sibs = shapeChildren(tree);
        const i = sibs.indexOf(s);
        if (how === 'front') tree.insertBefore(s, child(tree, 'p:extLst'));
        else if (how === 'back') tree.insertBefore(s, sibs[0]);
        else if (how === 'forward' && i < sibs.length - 1) tree.insertBefore(s, sibs[i + 1].nextSibling);
        else if (how === 'backward' && i > 0) tree.insertBefore(s, sibs[i - 1]);
      }
    });
  }

  /** Serialize shapes for the clipboard, including the relationship targets they reference. */
  copyShapes(slidePath: string, ids: string[]): ClipShape[] {
    return ids.map((id) => this.findShape(slidePath, id)).filter((s): s is Element => !!s).map((s) => {
      const rels: Record<string, { type: string; target: string }> = {};
      for (const el of [s, ...Array.from(s.getElementsByTagName('*'))]) {
        for (const a of Array.from(el.attributes)) {
          if (a.namespaceURI !== NS.r) continue;
          const r = this.pkg.rel(slidePath, a.value);
          if (r && !r.external) rels[a.value] = { type: r.type, target: r.target };
        }
      }
      return { xml: new XMLSerializer().serializeToString(s), rels };
    });
  }

  pasteShapes(slidePath: string, clip: ClipShape[], offset: number): string[] {
    return this.transact('붙이기', [slidePath, relsPathFor(slidePath)], () => {
      const ids: string[] = [];
      for (const c of clip) {
        const el = parseXml(c.xml).documentElement;
        for (const node of [el, ...Array.from(el.getElementsByTagName('*'))]) {
          for (const a of Array.from(node.attributes)) {
            if (a.namespaceURI !== NS.r) continue;
            const r = c.rels[a.value];
            if (r && this.pkg.has(r.target)) a.value = this.pkg.addRel(slidePath, r.type, r.target);
          }
        }
        const x = readXfrm(xfrmEl(el));
        if (x && offset) {
          const off = child(xfrmEl(el), 'a:off')!;
          off.setAttribute('x', String(x.x + offset));
          off.setAttribute('y', String(x.y + offset));
        }
        // Renumber nested ids so they stay unique within the slide.
        const doc = this.pkg.getXml(slidePath)!;
        let next = this.nextShapeId(doc);
        for (const c2 of [el, ...Array.from(el.getElementsByTagName('*'))]) {
          if (c2.localName === 'cNvPr' && c2.namespaceURI === NS.p) c2.setAttribute('id', String(next++));
        }
        const imported = doc.importNode(el, true) as Element;
        const tree = spTree(doc)!;
        tree.insertBefore(imported, child(tree, 'p:extLst'));
        ids.push(shapeId(imported));
      }
      return ids;
    });
  }

  // ---------------------------------------------------------------- formatting

  setFill(slidePath: string, ids: string[], hex: string | null) {
    this.transact('채우기', [slidePath], () => {
      for (const id of ids) {
        const s = this.findShape(slidePath, id);
        const pr = s && (child(s, 'p:spPr'));
        if (!pr) continue;
        for (const f of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']) child(pr, f)?.remove();
        const fill = hex ? parseXml(`<a:solidFill xmlns:a="${NS.a}"><a:srgbClr val="${hex}"/></a:solidFill>`).documentElement : create(pr.ownerDocument, 'a:noFill');
        const after = child(pr, 'a:prstGeom') ?? child(pr, 'a:custGeom') ?? child(pr, 'a:xfrm');
        pr.insertBefore(pr.ownerDocument.importNode(fill, true), after ? after.nextSibling : pr.firstChild);
      }
    });
  }

  setLine(slidePath: string, ids: string[], hex: string | null, widthPt?: number) {
    this.transact('선 모양', [slidePath], () => {
      for (const id of ids) {
        const s = this.findShape(slidePath, id);
        const pr = s && child(s, 'p:spPr');
        if (!pr) continue;
        const ln = ensure(pr, 'a:ln', ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']);
        if (widthPt != null) ln.setAttribute('w', String(Math.round(widthPt * EMU_PER_PT)));
        for (const f of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill']) child(ln, f)?.remove();
        const fill = hex ? parseXml(`<a:solidFill xmlns:a="${NS.a}"><a:srgbClr val="${hex}"/></a:solidFill>`).documentElement : create(pr.ownerDocument, 'a:noFill');
        ln.insertBefore(ln.ownerDocument.importNode(fill, true), ln.firstChild);
      }
    });
  }

  /** Apply character properties to every run of the given shapes (shape-level formatting). */
  applyRunProps(slidePath: string, ids: string[], props: RunProps) {
    this.transact('글자 모양', [slidePath], () => {
      for (const id of ids) {
        const s = this.findShape(slidePath, id);
        if (!s) continue;
        for (const tx of [...descendants(s, 'p:txBody'), ...descendants(s, 'a:txBody')]) {
          for (const p of children(tx, 'a:p')) {
            for (const r of [...children(p, 'a:r'), ...children(p, 'a:fld')]) setRunProps(ensure(r, 'a:rPr', ['a:t']), props);
            setRunProps(ensure(p, 'a:endParaRPr'), props);
          }
        }
      }
    });
  }

  applyParaProps(slidePath: string, ids: string[], props: { algn?: string; lvlDelta?: number; bullet?: 'none' | 'char' | 'num' }) {
    this.transact('문단 모양', [slidePath], () => {
      for (const id of ids) {
        const s = this.findShape(slidePath, id);
        if (!s) continue;
        for (const tx of [...descendants(s, 'p:txBody'), ...descendants(s, 'a:txBody')]) {
          for (const p of children(tx, 'a:p')) setParaProps(p, props);
        }
      }
    });
  }

  /** Replace a txBody's paragraphs (used when an in-place text edit is committed). */
  replaceParagraphs(slidePath: string, txBody: Element, paras: Element[], label = '텍스트 편집', extraParts: string[] = [], after?: () => void) {
    this.transact(label, [slidePath, ...extraParts], () => {
      for (const p of children(txBody, 'a:p')) p.remove();
      const ext = child(txBody, 'a:extLst');
      for (const p of paras) txBody.insertBefore(txBody.ownerDocument.importNode(p, true), ext);
      if (!children(txBody, 'a:p').length) txBody.insertBefore(create(txBody.ownerDocument, 'a:p'), ext);
      after?.();
    });
  }

  /** Find/replace inside single runs across all slides. Returns the number of replacements. */
  replaceAll(find: string, repl: string, matchCase: boolean): number {
    if (!find) return 0;
    const parts = this.slides().map((s) => s.path);
    let count = 0;
    this.transact('모두 바꾸기', parts, () => {
      const re = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
      for (const p of parts) {
        for (const t of descendants(this.pkg.getXml(p)!, 'a:t')) {
          const v = t.textContent || '';
          const n = v.match(re)?.length ?? 0;
          if (n) { t.textContent = v.replace(re, repl); count += n; }
        }
      }
    });
    return count;
  }

  /** Find occurrences: returns [{slideIndex, shapeId}] in order. */
  findAll(find: string, matchCase: boolean): { slide: number; shape: string }[] {
    const out: { slide: number; shape: string }[] = [];
    if (!find) return out;
    const f = matchCase ? find : find.toLowerCase();
    for (const s of this.slides()) {
      for (const sh of shapeChildren(spTree(this.pkg.getXml(s.path)))) {
        const text = descendants(sh, 'a:t').map((t) => t.textContent).join('\n');
        if ((matchCase ? text : text.toLowerCase()).includes(f)) out.push({ slide: s.index, shape: shapeId(sh) });
      }
    }
    return out;
  }
}

export interface ClipSlide { xml: string; rels: { id: string; type: string; target: string; external: boolean }[] }

export interface ClipShape { xml: string; rels: Record<string, { type: string; target: string }> }

export function escapeXml(s: string) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const RPR_FILL_BEFORE = ['a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];

export function setRunProps(rPr: Element, props: RunProps) {
  if (props.b != null) rPr.setAttribute('b', props.b ? '1' : '0');
  if (props.i != null) rPr.setAttribute('i', props.i ? '1' : '0');
  if (props.u != null) rPr.setAttribute('u', props.u ? 'sng' : 'none');
  if (props.strike != null) rPr.setAttribute('strike', props.strike ? 'sngStrike' : 'noStrike');
  if (props.sz != null) rPr.setAttribute('sz', String(Math.round(props.sz * 100)));
  if (props.color != null) {
    for (const f of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill']) child(rPr, f)?.remove();
    const fill = parseXml(`<a:solidFill xmlns:a="${NS.a}"><a:srgbClr val="${props.color}"/></a:solidFill>`).documentElement;
    let ref: Element | null = null;
    for (const b of RPR_FILL_BEFORE) { ref = child(rPr, b); if (ref) break; }
    rPr.insertBefore(rPr.ownerDocument.importNode(fill, true), ref);
  }
  if (props.font != null) {
    for (const tag of ['a:latin', 'a:ea']) {
      const before = tag === 'a:latin' ? ['a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'] : ['a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];
      ensure(rPr, tag, before).setAttribute('typeface', props.font);
    }
  }
}

export function setParaProps(p: Element, props: { algn?: string; lvlDelta?: number; bullet?: 'none' | 'char' | 'num' }) {
  const pPr = ensure(p, 'a:pPr');
  p.insertBefore(pPr, p.firstChild);
  if (props.algn) pPr.setAttribute('algn', props.algn);
  if (props.lvlDelta) {
    const lvl = Math.min(8, Math.max(0, num(pPr, 'lvl') + props.lvlDelta));
    if (lvl) pPr.setAttribute('lvl', String(lvl)); else pPr.removeAttribute('lvl');
  }
  if (props.bullet) {
    for (const b of ['a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip', 'a:buFont']) child(pPr, b)?.remove();
    const before = ['a:tabLst', 'a:defRPr', 'a:extLst'];
    if (props.bullet === 'none') ensure(pPr, 'a:buNone', before);
    else if (props.bullet === 'num') {
      ensure(pPr, 'a:buAutoNum', before).setAttribute('type', 'arabicPeriod');
      if (!pPr.hasAttribute('marL')) { pPr.setAttribute('marL', '342900'); pPr.setAttribute('indent', '-342900'); }
    } else {
      ensure(pPr, 'a:buFont', before).setAttribute('typeface', 'Arial');
      ensure(pPr, 'a:buChar', before).setAttribute('char', '•');
      if (!pPr.hasAttribute('marL')) { pPr.setAttribute('marL', '285750'); pPr.setAttribute('indent', '-285750'); }
    }
  }
}

export { dirname };
