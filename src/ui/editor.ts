// Interactive slide editor: selection, move/resize/rotate, shape drawing, in-place text editing.
import { PptxDocument, RunProps, Xfrm, setParaProps, xfrmEl } from '../engine/document';
import { ImageCache, RenderedShape, RenderedSlide, TextTarget, renderSlide } from '../engine/render';
import { domToParagraphs } from '../engine/text';
import { child } from '../engine/xml';

const EMU = 12700; // per CSS px at zoom 1

export type InsertMode = { kind: 'shape'; prst: string; label: string } | { kind: 'textbox' } | null;

export interface EditorHost {
  doc(): PptxDocument | null;
  images(): ImageCache;
  selectionChanged(): void;
  status(msg: string): void;
  modeChanged(): void;
}

interface EditState {
  shape: RenderedShape;
  target: TextTarget;
  box: HTMLElement; // element whose rect is shown as the editing frame
  startHtml: string;
}

interface Drag {
  kind: 'move' | 'resize' | 'rotate' | 'rubber' | 'draw';
  handle?: string;
  startX: number; startY: number; // slide EMU
  start: Map<string, Xfrm>;
  moved: boolean;
  shift: boolean;
  preview?: HTMLElement;
  last?: Map<string, Xfrm>;
}

export class SlideEditor {
  rendered: RenderedSlide | null = null;
  slidePath = '';
  zoom = 1;
  selection: string[] = [];
  editing: EditState | null = null;
  insertMode: InsertMode = null;
  showGrid = false;
  private overlay: HTMLElement;
  private drag: Drag | null = null;
  private savedRange: Range | null = null;

  constructor(private host: EditorHost, private slideHost: HTMLElement) {
    this.overlay = document.createElement('div');
    this.overlay.id = 'overlay';
    slideHost.addEventListener('pointerdown', (e) => this.onPointerDown(e));
    slideHost.addEventListener('dblclick', (e) => this.onDblClick(e));
    window.addEventListener('pointermove', (e) => this.onPointerMove(e));
    window.addEventListener('pointerup', (e) => this.onPointerUp(e));
    document.addEventListener('selectionchange', () => { if (this.editing) this.host.selectionChanged(); });
  }

  // ------------------------------------------------------------------ rendering

  render(slidePath: string) {
    const doc = this.host.doc();
    if (this.editing && slidePath !== this.slidePath) this.commitEdit();
    this.editing = null;
    this.slidePath = slidePath;
    this.slideHost.replaceChildren();
    if (!doc || !slidePath) { this.rendered = null; this.syncOverlay(); return; }
    this.rendered = renderSlide(doc, slidePath, this.host.images(), { editor: true });
    this.slideHost.appendChild(this.rendered.el);
    if (this.showGrid) {
      const g = document.createElement('div');
      g.className = 'grid-layer';
      this.rendered.el.appendChild(g);
    }
    this.slideHost.appendChild(this.overlay);
    const ids = new Set(this.rendered.shapes.map((s) => s.id));
    this.selection = this.selection.filter((id) => ids.has(id));
    this.applyZoom();
  }

  setZoom(z: number) {
    this.zoom = Math.min(4, Math.max(0.1, z));
    this.applyZoom();
  }

  private applyZoom() {
    if (!this.rendered) return;
    this.rendered.el.style.transform = `scale(${this.zoom})`;
    this.slideHost.style.width = this.rendered.width * this.zoom + 'px';
    this.slideHost.style.height = this.rendered.height * this.zoom + 'px';
    this.syncOverlay();
  }

  shape(id: string) { return this.rendered?.shapes.find((s) => s.id === id) ?? null; }
  selectedShapes() { return this.selection.map((id) => this.shape(id)).filter((s): s is RenderedShape => !!s); }

  select(ids: string[]) {
    if (this.editing) this.commitEdit();
    this.selection = ids;
    this.syncOverlay();
    this.host.selectionChanged();
  }

  // ------------------------------------------------------------------ overlay

  syncOverlay() {
    const o = this.overlay;
    o.replaceChildren();
    if (!this.rendered) return;
    const z = this.zoom;
    const multi = this.selection.length > 1;
    for (const s of this.selectedShapes()) {
      const x = this.liveXfrm(s);
      const box = document.createElement('div');
      box.className = 'sel-box' + (multi ? ' multi' : '') + (this.editing?.shape === s ? ' editing' : '');
      let h = x.h;
      if (s.el.classList.contains('autofit')) h = Math.max(h, (s.text?.rendered.root.offsetHeight ?? 0) * EMU);
      box.style.cssText = `left:${(x.x / EMU) * z}px;top:${(x.y / EMU) * z}px;width:${(x.w / EMU) * z}px;height:${(h / EMU) * z}px;` + (x.rot ? `transform:rotate(${x.rot}deg)` : '');
      if (this.editing?.target.cell && this.editing.shape === s) {
        const td = this.editing.box;
        box.style.borderStyle = 'solid';
        const cell = document.createElement('div');
        cell.className = 'sel-box editing';
        cell.style.cssText = `left:${td.offsetLeft * z}px;top:${td.offsetTop * z}px;width:${td.offsetWidth * z}px;height:${td.offsetHeight * z}px`;
        box.appendChild(cell);
      }
      if (!multi && !s.locked && !this.editing) {
        const hs = s.isLine ? ['nw', 'se'] : ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
        for (const hd of hs) {
          const el = document.createElement('div');
          el.className = 'handle';
          el.dataset.h = hd;
          el.style.left = (hd.includes('w') ? 0 : hd.includes('e') ? 100 : 50) + '%';
          el.style.top = (hd.includes('n') ? 0 : hd.includes('s') ? 100 : 50) + '%';
          box.appendChild(el);
        }
        if (!s.isLine) {
          const stem = document.createElement('div');
          stem.className = 'rot-stem';
          stem.style.cssText = 'left:50%;top:-22px;height:22px';
          box.appendChild(stem);
          const r = document.createElement('div');
          r.className = 'handle rot';
          r.dataset.h = 'rot';
          r.style.left = '50%';
          r.style.top = '-22px';
          box.appendChild(r);
        }
      }
      o.appendChild(box);
    }
  }

  private liveXfrm(s: RenderedShape): Xfrm {
    return this.drag?.last?.get(s.id) ?? s.xfrm;
  }

  // ------------------------------------------------------------------ pointer

  private toSlide(e: { clientX: number; clientY: number }) {
    const r = this.slideHost.getBoundingClientRect();
    return { x: ((e.clientX - r.left) / this.zoom) * EMU, y: ((e.clientY - r.top) / this.zoom) * EMU };
  }

  private hitShape(target: EventTarget | null): RenderedShape | null {
    const el = (target as HTMLElement | null)?.closest?.('.shape.top:not(.decor)') as HTMLElement | null;
    return el ? this.shape(el.dataset.id!) : null;
  }

  private onPointerDown(e: PointerEvent) {
    if (e.button !== 0 || !this.rendered) return;
    const t = e.target as HTMLElement;
    const p = this.toSlide(e);
    if (this.insertMode) {
      e.preventDefault();
      const preview = document.createElement('div');
      preview.className = 'draw-preview';
      this.overlay.appendChild(preview);
      this.drag = { kind: 'draw', startX: p.x, startY: p.y, start: new Map(), moved: false, shift: e.shiftKey, preview };
      return;
    }
    if (this.editing && this.editing.box.contains(t)) return; // native caret placement
    if (this.editing) this.commitEdit();

    const handle = t.closest('.handle') as HTMLElement | null;
    if (handle) {
      e.preventDefault();
      const s = this.selectedShapes()[0];
      if (!s) return;
      this.drag = { kind: handle.dataset.h === 'rot' ? 'rotate' : 'resize', handle: handle.dataset.h, startX: p.x, startY: p.y, start: new Map([[s.id, s.xfrm]]), moved: false, shift: e.shiftKey };
      return;
    }
    const s = this.hitShape(t);
    (document.getElementById('scroll-container') as HTMLElement).focus({ preventScroll: true });
    if (s) {
      e.preventDefault();
      if (e.shiftKey || e.metaKey || e.ctrlKey) {
        const has = this.selection.includes(s.id);
        this.select(has ? this.selection.filter((x) => x !== s.id) : [...this.selection, s.id]);
        if (has) return;
      } else if (!this.selection.includes(s.id)) this.select([s.id]);
      const movable = this.selectedShapes().filter((x) => !x.locked);
      this.drag = { kind: 'move', startX: p.x, startY: p.y, start: new Map(movable.map((x) => [x.id, x.xfrm])), moved: false, shift: false };
      return;
    }
    e.preventDefault();
    if (!e.shiftKey) this.select([]);
    const preview = document.createElement('div');
    preview.className = 'rubber';
    this.overlay.appendChild(preview);
    this.drag = { kind: 'rubber', startX: p.x, startY: p.y, start: new Map(), moved: false, shift: e.shiftKey, preview };
  }

  private onPointerMove(e: PointerEvent) {
    const d = this.drag;
    if (!d || !this.rendered) return;
    const p = this.toSlide(e);
    let dx = p.x - d.startX, dy = p.y - d.startY;
    if (!d.moved && Math.hypot(dx, dy) < (3 / this.zoom) * EMU) return;
    d.moved = true;
    const z = this.zoom;
    const rectCss = (x: number, y: number, w: number, h: number) => `left:${(x / EMU) * z}px;top:${(y / EMU) * z}px;width:${(w / EMU) * z}px;height:${(h / EMU) * z}px`;

    if (d.kind === 'rubber' || d.kind === 'draw') {
      let w = dx, h = dy;
      if (d.kind === 'draw' && e.shiftKey) { const m = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * m; h = Math.sign(h || 1) * m; }
      d.preview!.style.cssText = rectCss(Math.min(d.startX, d.startX + w), Math.min(d.startY, d.startY + h), Math.abs(w), Math.abs(h));
      d.last = new Map([['', { x: d.startX, y: d.startY, w, h, rot: 0, flipH: false, flipV: false }]]);
      return;
    }

    const next = new Map<string, Xfrm>();
    this.clearGuides();
    if (d.kind === 'move') {
      if (e.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      if (!e.altKey) [dx, dy] = this.snapMove(d.start, dx, dy);
      for (const [id, x] of d.start) next.set(id, { ...x, x: x.x + dx, y: x.y + dy });
    } else if (d.kind === 'resize') {
      const [id, x] = [...d.start][0];
      const s = this.shape(id)!;
      next.set(id, resizeXfrm(x, d.handle!, dx, dy, e.shiftKey || (s.kind === 'pic' && d.handle!.length === 2)));
    } else if (d.kind === 'rotate') {
      const [id, x] = [...d.start][0];
      const cx = x.x + x.w / 2, cy = x.y + x.h / 2;
      let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (e.shiftKey) ang = Math.round(ang / 15) * 15;
      ang = ((ang % 360) + 360) % 360;
      next.set(id, { ...x, rot: Math.round(ang * 10) / 10 });
    }
    d.last = next;
    for (const [id, x] of next) {
      const s = this.shape(id);
      if (!s) continue;
      s.el.style.left = x.x / EMU + 'px';
      s.el.style.top = x.y / EMU + 'px';
      s.el.style.width = x.w / EMU + 'px';
      s.el.style.height = x.h / EMU + 'px';
      s.el.style.transform = x.rot ? `rotate(${x.rot}deg)` : '';
    }
    this.syncOverlay();
    if (d.kind === 'move') this.drawGuides();
    const f = [...next.values()][0];
    if (f) this.host.status(`X ${(f.x / 360000).toFixed(2)}cm  Y ${(f.y / 360000).toFixed(2)}cm  W ${(f.w / 360000).toFixed(2)}cm  H ${(f.h / 360000).toFixed(2)}cm${f.rot ? `  ${f.rot}°` : ''}`);
  }

  private onPointerUp(e: PointerEvent) {
    const d = this.drag;
    if (!d) return;
    this.drag = null;
    this.clearGuides();
    d.preview?.remove();
    const doc = this.host.doc();
    if (!doc || !this.rendered) return;

    if (d.kind === 'draw') {
      const mode = this.insertMode!;
      const r = d.last?.get('');
      const p = this.toSlide(e);
      let x: Xfrm;
      if (!d.moved || !r) {
        const w = mode.kind === 'textbox' ? 6 * 360000 : 2.54 * 360000;
        const h = mode.kind === 'textbox' ? 0.9 * 360000 : mode.kind === 'shape' && mode.prst === 'line' ? 0 : 2.54 * 360000;
        x = { x: p.x, y: p.y, w, h, rot: 0, flipH: false, flipV: false };
      } else {
        x = { x: Math.min(r.x, r.x + r.w), y: Math.min(r.y, r.y + r.h), w: Math.abs(r.w), h: Math.abs(r.h), rot: 0, flipH: r.w < 0, flipV: r.h < 0 };
      }
      this.insertMode = null;
      this.rendered.el.classList.remove('inserting');
      this.host.modeChanged();
      let id: string;
      if (mode.kind === 'textbox') id = doc.addTextBox(this.slidePath, { ...x, h: Math.max(x.h, 0.9 * 360000) });
      else id = doc.addShape(this.slidePath, mode.prst, x, mode.label);
      this.select([id]);
      if (mode.kind === 'textbox') this.enterEdit(id);
      return;
    }
    if (d.kind === 'rubber') {
      const r = d.last?.get('');
      if (!r) return;
      const x0 = Math.min(r.x, r.x + r.w), y0 = Math.min(r.y, r.y + r.h), x1 = Math.max(r.x, r.x + r.w), y1 = Math.max(r.y, r.y + r.h);
      const hits = this.rendered.shapes.filter((s) => s.xfrm.x >= x0 && s.xfrm.y >= y0 && s.xfrm.x + s.xfrm.w <= x1 && s.xfrm.y + s.xfrm.h <= y1).map((s) => s.id);
      this.select(d.shift ? [...new Set([...this.selection, ...hits])] : hits);
      return;
    }
    if (!d.moved || !d.last) return;
    const label = d.kind === 'move' ? '개체 이동' : d.kind === 'resize' ? '크기 조절' : '회전';
    doc.setXfrms(this.slidePath, [...d.last].map(([id, x]) => ({ id, x })), label);
  }

  private onDblClick(e: MouseEvent) {
    if (this.insertMode) return;
    const s = this.hitShape(e.target);
    if (!s || s.locked) return;
    const td = (e.target as HTMLElement).closest('td[data-cell]') as HTMLElement | null;
    if (td) {
      const [r, c] = td.dataset.cell!.split(',').map(Number);
      this.enterEdit(s.id, { r, c }, e);
    } else if (s.text) this.enterEdit(s.id, undefined, e);
  }

  // ------------------------------------------------------------------ snapping

  private guides: { v: number[]; h: number[] } = { v: [], h: [] };

  private snapMove(start: Map<string, Xfrm>, dx: number, dy: number): [number, number] {
    const doc = this.host.doc()!;
    const { cx, cy } = doc.slideSize;
    const boxes = [...start.values()];
    const x0 = Math.min(...boxes.map((b) => b.x)) + dx, x1 = Math.max(...boxes.map((b) => b.x + b.w)) + dx;
    const y0 = Math.min(...boxes.map((b) => b.y)) + dy, y1 = Math.max(...boxes.map((b) => b.y + b.h)) + dy;
    const vx = [0, cx / 2, cx], hy = [0, cy / 2, cy];
    for (const s of this.rendered!.shapes) {
      if (start.has(s.id)) continue;
      vx.push(s.xfrm.x, s.xfrm.x + s.xfrm.w / 2, s.xfrm.x + s.xfrm.w);
      hy.push(s.xfrm.y, s.xfrm.y + s.xfrm.h / 2, s.xfrm.y + s.xfrm.h);
    }
    const tol = (6 / this.zoom) * EMU;
    const best = (vals: number[], lines: number[]) => {
      let bd = Infinity, line = 0;
      for (const v of vals) for (const l of lines) { const dd = l - v; if (Math.abs(dd) < Math.abs(bd)) { bd = dd; line = l; } }
      return Math.abs(bd) <= tol ? { d: bd, line } : null;
    };
    const bx = best([x0, (x0 + x1) / 2, x1], vx);
    const by = best([y0, (y0 + y1) / 2, y1], hy);
    this.guides = { v: bx ? [bx.line] : [], h: by ? [by.line] : [] };
    return [dx + (bx?.d ?? 0), dy + (by?.d ?? 0)];
  }

  private drawGuides() {
    const z = this.zoom;
    const { width, height } = this.rendered!;
    for (const v of this.guides.v) {
      const g = document.createElement('div');
      g.className = 'guide';
      g.style.cssText = `left:${(v / EMU) * z}px;top:0;width:1px;height:${height * z}px`;
      this.overlay.appendChild(g);
    }
    for (const h of this.guides.h) {
      const g = document.createElement('div');
      g.className = 'guide';
      g.style.cssText = `top:${(h / EMU) * z}px;left:0;height:1px;width:${width * z}px`;
      this.overlay.appendChild(g);
    }
  }

  private clearGuides() {
    this.overlay.querySelectorAll('.guide').forEach((g) => g.remove());
  }

  // ------------------------------------------------------------------ insert mode

  startInsert(mode: InsertMode) {
    if (this.editing) this.commitEdit();
    this.insertMode = mode;
    this.rendered?.el.classList.toggle('inserting', !!mode);
    this.host.modeChanged();
    this.host.status(mode ? '슬라이드에서 끌어서 그리거나 클릭하세요. (Esc: 취소)' : '');
  }

  // ------------------------------------------------------------------ text editing

  enterEdit(id: string, cell?: { r: number; c: number }, at?: { clientX: number; clientY: number }) {
    const s = this.shape(id);
    if (!s || s.locked) return;
    const target = cell ? s.cells.find((c) => c.cell!.r === cell.r && c.cell!.c === cell.c) : s.text ?? s.cells[0];
    if (!target) return;
    this.selection = [id];
    const r = target.rendered;
    if (r.root.classList.contains('ph-prompt')) {
      // Swap the greyed layout prompt for the real (empty) body.
      r.root.classList.remove('ph-prompt');
      const p = document.createElement('p');
      p.className = 'tx-p';
      p.dataset.pi = '0';
      const first = r.content.querySelector('p') as HTMLElement | null;
      if (first) { p.style.cssText = first.style.cssText; p.className = first.className; p.style.removeProperty('--bu'); p.classList.remove('has-bullet'); }
      p.appendChild(document.createElement('br'));
      r.content.replaceChildren(p);
    }
    const box = target.cell ? (r.root.closest('td') as HTMLElement) : s.el;
    box.classList.add('editing');
    s.el.classList.add('editing');
    r.content.contentEditable = 'true';
    r.content.spellcheck = false;
    this.editing = { shape: s, target, box, startHtml: r.content.innerHTML };
    document.execCommand('styleWithCSS', false, 'true');
    r.content.focus({ preventScroll: true });
    const sel = window.getSelection()!;
    let range: Range | null = null;
    if (at) {
      const d = document as Document & { caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null };
      if (document.caretRangeFromPoint) range = document.caretRangeFromPoint(at.clientX, at.clientY);
      else if (d.caretPositionFromPoint) {
        const cp = d.caretPositionFromPoint(at.clientX, at.clientY);
        if (cp) { range = document.createRange(); range.setStart(cp.offsetNode, cp.offset); }
      }
      if (range && !r.content.contains(range.startContainer)) range = null;
    }
    if (!range) {
      range = document.createRange();
      range.selectNodeContents(r.content);
      range.collapse(false);
    }
    sel.removeAllRanges();
    sel.addRange(range);
    r.content.addEventListener('input', this.onInput);
    r.content.addEventListener('paste', this.onPaste);
    r.content.addEventListener('keydown', this.onEditKey);
    this.syncOverlay();
    this.host.modeChanged();
    this.host.selectionChanged();
  }

  private onInput = () => this.syncOverlay();

  private onPaste = (e: ClipboardEvent) => {
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    document.execCommand('insertText', false, text);
  };

  private onEditKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); this.commitEdit(); }
    else if (e.key === 'Tab' && !this.editing?.target.cell) {
      e.preventDefault(); e.stopPropagation();
      this.paraOp({ lvlDelta: e.shiftKey ? -1 : 1 });
    }
  };

  /** Commit the current text edit to the model. Returns true if something changed. */
  commitEdit(): boolean {
    const ed = this.editing;
    if (!ed) return false;
    this.editing = null;
    const doc = this.host.doc()!;
    const { rendered, txBody } = ed.target;
    rendered.content.removeEventListener('input', this.onInput);
    rendered.content.removeEventListener('paste', this.onPaste);
    rendered.content.removeEventListener('keydown', this.onEditKey);
    rendered.content.contentEditable = 'false';
    ed.box.classList.remove('editing');
    ed.shape.el.classList.remove('editing');
    window.getSelection()?.removeAllRanges();
    const changed = rendered.content.innerHTML !== ed.startHtml;
    if (changed) {
      const paras = domToParagraphs(rendered.content, rendered.paras, txBody.ownerDocument as XMLDocument, rendered.layout.fontScale);
      const shapeEl = ed.shape.shape;
      const autoH = ed.shape.el.classList.contains('autofit') ? rendered.root.offsetHeight * EMU : 0;
      doc.replaceParagraphs(this.slidePath, txBody, paras, '텍스트 편집', [], () => {
        const ext = child(xfrmEl(shapeEl), 'a:ext');
        if (autoH && ext) ext.setAttribute('cy', String(Math.round(autoH)));
      });
    } else this.render(this.slidePath);
    this.host.modeChanged();
    return changed;
  }

  cancelEdit() {
    if (!this.editing) return;
    this.editing = null;
    this.render(this.slidePath);
    this.host.modeChanged();
  }

  // ------------------------------------------------------------------ formatting

  /** Paragraph indices (data-pi order) touched by the current DOM selection. */
  private selectedParaIndices(): number[] {
    const ed = this.editing!;
    const blocks = Array.from(ed.target.rendered.content.children) as HTMLElement[];
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return blocks.map((_, i) => i);
    const range = sel.getRangeAt(0);
    const out = blocks.map((b, i) => (range.intersectsNode(b) ? i : -1)).filter((i) => i >= 0);
    return out.length ? out : [0];
  }

  rememberSelection() {
    const sel = window.getSelection();
    if (this.editing && sel?.rangeCount) this.savedRange = sel.getRangeAt(0).cloneRange();
  }

  private restoreSelection() {
    if (!this.editing) return;
    this.editing.target.rendered.content.focus({ preventScroll: true });
    if (this.savedRange) {
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(this.savedRange);
      this.savedRange = null;
    }
  }

  applyCharProps(props: RunProps) {
    const doc = this.host.doc();
    if (!doc) return;
    if (this.editing) {
      this.restoreSelection();
      const content = this.editing.target.rendered.content;
      const scale = this.editing.target.rendered.layout.fontScale;
      const cur = this.caretStyle();
      if (props.b != null && props.b !== cur?.b) document.execCommand('bold');
      if (props.i != null && props.i !== cur?.i) document.execCommand('italic');
      if (props.u != null && props.u !== cur?.u) document.execCommand('underline');
      if (props.strike != null && props.strike !== cur?.strike) document.execCommand('strikeThrough');
      if (props.color) document.execCommand('foreColor', false, '#' + props.color);
      if (props.font) document.execCommand('fontName', false, props.font);
      if (props.sz) {
        document.execCommand('fontSize', false, '7');
        content.querySelectorAll<HTMLElement>('font[size="7"], span[style*="xxx-large"]').forEach((el) => {
          const span = document.createElement('span');
          span.style.fontSize = props.sz! * scale + 'px';
          if (el.tagName === 'SPAN') { el.style.fontSize = props.sz! * scale + 'px'; return; }
          while (el.firstChild) span.appendChild(el.firstChild);
          el.replaceWith(span);
        });
      }
      this.syncOverlay();
      this.host.selectionChanged();
      return;
    }
    const ids = this.selectedShapes().filter((s) => !s.locked && (s.text || s.cells.length)).map((s) => s.id);
    if (ids.length) doc.applyRunProps(this.slidePath, ids, props);
  }

  align(algn: 'l' | 'ctr' | 'r' | 'just') {
    const doc = this.host.doc();
    if (!doc) return;
    if (this.editing) {
      const blocks = Array.from(this.editing.target.rendered.content.children) as HTMLElement[];
      for (const i of this.selectedParaIndices()) {
        blocks[i].dataset.algn = algn;
        blocks[i].style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[algn];
      }
      this.host.selectionChanged();
      return;
    }
    const ids = this.selectedShapes().filter((s) => !s.locked).map((s) => s.id);
    if (ids.length) doc.applyParaProps(this.slidePath, ids, { algn });
  }

  /** Bullet/level changes. While editing, commits and re-enters so the model renders the result. */
  paraOp(props: { lvlDelta?: number; bullet?: 'none' | 'char' | 'num' }) {
    const doc = this.host.doc();
    if (!doc) return;
    if (this.editing) {
      const idxs = this.selectedParaIndices();
      const id = this.editing.shape.id;
      const cell = this.editing.target.cell;
      this.commitEdit();
      const s = this.shape(id);
      const target = cell ? s?.cells.find((c) => c.cell!.r === cell.r && c.cell!.c === cell.c) : s?.text;
      if (!target) return;
      const tx = target.txBody;
      doc.transact('문단 모양', [this.slidePath], () => {
        const paras = Array.from(tx.children).filter((c) => c.localName === 'p');
        for (const i of idxs) if (paras[i]) setParaProps(paras[i], props);
      });
      this.enterEdit(id, cell);
      return;
    }
    const ids = this.selectedShapes().filter((s) => !s.locked).map((s) => s.id);
    if (ids.length) doc.applyParaProps(this.slidePath, ids, props);
  }

  /** Style at the caret (editing) or of the first selected shape's first run. */
  caretStyle(): { font: string; size: number; b: boolean; i: boolean; u: boolean; strike: boolean; color: string; algn: string } | null {
    let el: Element | null = null;
    let scale = 1;
    let block: HTMLElement | null = null;
    if (this.editing) {
      const sel = window.getSelection();
      const n = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
      const content = this.editing.target.rendered.content;
      if (n && content.contains(n)) {
        el = n.nodeType === Node.TEXT_NODE ? n.parentElement : (n as Element);
        if (el === content) el = content.querySelector('span') ?? content.querySelector('p');
        else if (el && el.tagName === 'P' && el.firstElementChild?.tagName === 'SPAN') el = el.firstElementChild;
      }
      scale = this.editing.target.rendered.layout.fontScale;
    } else {
      const s = this.selectedShapes()[0];
      const r = s?.text?.rendered ?? s?.cells[0]?.rendered;
      if (r) { el = r.content.querySelector('span') ?? r.content.querySelector('p'); scale = r.layout.fontScale; }
    }
    if (!el) return null;
    block = el.closest('p');
    const cs = getComputedStyle(el);
    let u = false, strike = false;
    for (let e: Element | null = el; e && e.tagName !== 'DIV'; e = e.parentElement) {
      const l = getComputedStyle(e).textDecorationLine;
      if (l.includes('underline') || e.tagName === 'U') u = true;
      if (l.includes('line-through') || e.tagName === 'S' || e.tagName === 'STRIKE') strike = true;
    }
    const ta = block ? getComputedStyle(block).textAlign : 'left';
    return {
      font: cs.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
      size: Math.round((parseFloat(cs.fontSize) / scale) * 10) / 10,
      b: Number(cs.fontWeight) >= 600 || cs.fontWeight === 'bold',
      i: cs.fontStyle === 'italic',
      u, strike,
      color: cs.color,
      algn: ta === 'center' ? 'ctr' : ta === 'right' || ta === 'end' ? 'r' : ta === 'justify' ? 'just' : 'l',
    };
  }
}

export function resizeXfrm(x: Xfrm, handle: string, dx: number, dy: number, keepAspect: boolean): Xfrm {
  const rad = (x.rot * Math.PI) / 180;
  const cos = Math.cos(rad), sin = Math.sin(rad);
  // Pointer delta in the shape's local (unrotated) frame.
  const lx = dx * cos + dy * sin;
  const ly = -dx * sin + dy * cos;
  let w = x.w, h = x.h, cxl = 0, cyl = 0;
  if (handle.includes('e')) { w = x.w + lx; }
  if (handle.includes('w')) { w = x.w - lx; }
  if (handle.includes('s')) { h = x.h + ly; }
  if (handle.includes('n')) { h = x.h - ly; }
  const min = 12700 * 2;
  w = Math.max(min, w);
  h = Math.max(x.h === 0 ? 0 : min, h);
  if (keepAspect && x.w > 0 && x.h > 0) {
    if (handle.length === 2) {
      const k = Math.max(w / x.w, h / x.h);
      w = x.w * k; h = x.h * k;
    } else if (handle === 'e' || handle === 'w') h = (x.h * w) / x.w;
    else w = (x.w * h) / x.h;
  }
  const dw = w - x.w, dh = h - x.h;
  cxl = handle.includes('w') ? -dw / 2 : handle.includes('e') ? dw / 2 : 0;
  cyl = handle.includes('n') ? -dh / 2 : handle.includes('s') ? dh / 2 : 0;
  const cx = x.x + x.w / 2 + cxl * cos - cyl * sin;
  const cy = x.y + x.h / 2 + cxl * sin + cyl * cos;
  return { ...x, x: cx - w / 2, y: cy - h / 2, w, h };
}
