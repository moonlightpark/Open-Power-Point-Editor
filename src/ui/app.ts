// Application controller: document lifecycle, command registration, UI state sync.
import { ClipShape, ClipSlide, PptxDocument, RunProps, Xfrm } from '../engine/document';
import { ImageCache } from '../engine/render';
import { Commands } from './commands';
import { closeContextMenu, closePopover, colorPopover, contextMenu, dialog, html, layoutDialog, objectPropsDialog, shapePopover, tablePopover, toast } from './dialogs';
import { SlideEditor } from './editor';
import { installIcons } from './icons';
import { drawRulers } from './ruler';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { isDesktop, pickImage, pickOpenFile, saveFile } from './shell';
import { printSlides, startSlideShow } from './slideshow';
import { Thumbnails } from './thumbnails';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const FONTS = ['맑은 고딕', 'Apple SD Gothic Neo', '나눔고딕', '나눔명조', '굴림', '돋움', '바탕', '궁서', 'Noto Sans KR', 'Pretendard', 'Arial', 'Calibri', 'Calibri Light', 'Helvetica', 'Times New Roman', 'Georgia', 'Verdana', 'Tahoma', 'Segoe UI', 'Consolas', 'Courier New'];
const ZOOMS = [0.1, 0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.25, 1.5, 2, 3, 4];

export class App {
  doc: PptxDocument | null = null;
  images!: ImageCache;
  fileName = '새 프레젠테이션.pptx';
  fileHandle: FileSystemFileHandle | undefined;
  filePath: string | undefined; // desktop shell
  current = 0;
  slideSel = new Set<number>();
  fitZoom = true;
  showRuler = true;
  showPanel = true;
  showNotes = false;
  clipboard: { items: ClipShape[]; text: string; pasteCount: number } | null = null;
  slideClipboard: ClipSlide[] | null = null;
  lastTextColor = 'C00000';
  readonly cmd = new Commands();
  readonly editor: SlideEditor;
  readonly thumbs: Thumbnails;
  private slidePaths: string[] = [];

  constructor() {
    this.editor = new SlideEditor({
      doc: () => this.doc,
      images: () => this.images,
      selectionChanged: () => { this.syncStyleBar(); this.syncRulers(); },
      status: (m) => this.message(m),
      modeChanged: () => this.syncStatus(),
    }, $('slide-host'));
    this.thumbs = new Thumbnails({
      doc: () => this.doc,
      images: () => this.images,
      current: () => this.current,
      selectedSlides: () => this.slideSel,
      goto: (i, ext) => this.gotoSlide(i, ext),
      move: (from, to) => this.moveSlides(from, to),
      contextMenu: (i, x, y) => this.slideContextMenu(i, x, y),
      action: (cmd, i) => {
        this.thumbs.focusList();
        if (cmd === 'new-after') { this.insertNewSlide(i + 1); return; }
        // Hover buttons act on the hovered slide, or on the whole selection if it is part of it.
        if (!this.slideSel.has(i)) this.gotoSlide(i);
        this.cmd.run(cmd === 'duplicate' ? 'slide:duplicate' : 'slide:delete');
      },
    }, $('slide-list'));
    this.cmd.onError = (e) => { console.error(e); toast('오류: ' + ((e as Error).message ?? e)); };
    this.registerCommands();
    this.bindUi();
  }

  // ------------------------------------------------------------------ document lifecycle

  async open(bytes: Uint8Array, name: string, handle?: FileSystemFileHandle, path?: string) {
    const doc = await PptxDocument.load(bytes);
    this.images?.dispose();
    this.doc = doc;
    this.images = new ImageCache(doc);
    this.fileName = name;
    this.fileHandle = handle;
    this.filePath = path;
    this.current = 0;
    this.slideSel = new Set([0]);
    this.editor.selection = [];
    doc.onChange = () => this.onDocChange();
    await Promise.all(doc.slides().map((s) => this.images.ensureForSlide(s.path)));
    this.thumbs.clear();
    this.slidePaths = doc.slides().map((s) => s.path);
    this.thumbs.update();
    this.renderCurrent();
    this.zoomFit();
    this.updateTitle();
    this.message(`${name} 열기 완료 — 슬라이드 ${doc.slides().length}장`);
  }

  async newDocument() {
    const res = await fetch('./blank.pptx');
    const bytes = new Uint8Array(await res.arrayBuffer());
    await this.open(bytes, '새 프레젠테이션.pptx');
    const doc = this.doc!;
    const title = doc.layouts()[0];
    if (title) doc.addSlide(title.path, 0);
    doc.dirty = false;
    this.updateTitle();
    this.message('새 프레젠테이션');
  }

  private onDocChange() {
    const doc = this.doc!;
    const paths = doc.slides().map((s) => s.path);
    const structural = paths.join('|') !== this.slidePaths.join('|');
    this.slidePaths = paths;
    if (this.current >= paths.length) this.current = Math.max(0, paths.length - 1);
    const cur = paths[this.current];
    // New images (insert/paste) need loading before the render can show them.
    this.images.ensureForSlide(cur).then(() => {
      if (this.editor.editing) return;
      this.renderCurrent();
      this.thumbs.update(structural ? undefined : [cur]);
    });
    if (!this.editor.editing) this.renderCurrent();
    this.thumbs.update(structural ? undefined : [cur]);
    this.updateTitle();
  }

  private renderCurrent() {
    const s = this.doc?.slides()[this.current];
    this.editor.render(s?.path ?? '');
    this.syncStatus();
    this.syncStyleBar();
    this.syncRulers();
    $('notes-text').textContent = s ? this.doc!.notesText(s.path) || '(노트 없음)' : '';
  }

  gotoSlide(i: number, extend?: 'toggle' | 'range') {
    if (!this.doc) return;
    const n = this.doc.slides().length;
    if (!n) return;
    i = Math.max(0, Math.min(n - 1, i));
    if (this.editor.editing) this.editor.commitEdit();
    if (extend === 'toggle') { if (this.slideSel.has(i) && this.slideSel.size > 1) this.slideSel.delete(i); else this.slideSel.add(i); }
    else if (extend === 'range') { const [a, b] = [Math.min(this.current, i), Math.max(this.current, i)]; this.slideSel = new Set(Array.from({ length: b - a + 1 }, (_, k) => a + k)); }
    else this.slideSel = new Set([i]);
    if (i !== this.current) { this.current = i; this.editor.selection = []; this.renderCurrent(); }
    this.thumbs.markCurrent();
  }

  private moveSlides(from: number, to: number) {
    this.doc!.moveSlide(from, to);
    this.current = to;
    this.slideSel = new Set([to]);
    this.onDocChange();
  }

  private updateTitle() {
    document.title = `${this.doc?.dirty ? '● ' : ''}${this.fileName} — Opptx`;
    if (isDesktop()) getCurrentWindow().setTitle(document.title).catch(() => {});
  }

  message(m: string) { $('sb-message').textContent = m; }

  // ------------------------------------------------------------------ zoom

  setZoom(z: number) {
    this.fitZoom = false;
    this.editor.setZoom(z);
    this.syncStatus();
    this.syncRulers();
  }

  zoomFit() {
    if (!this.doc) return;
    this.fitZoom = true;
    const sc = $('scroll-container');
    const { cx, cy } = this.doc.slideSize;
    const z = Math.min((sc.clientWidth - 80) / (cx / 12700), (sc.clientHeight - 80) / (cy / 12700));
    this.editor.setZoom(Math.max(0.1, z));
    this.syncStatus();
    this.syncRulers();
  }

  private stepZoom(dir: 1 | -1) {
    const z = this.editor.zoom;
    const next = dir > 0 ? ZOOMS.find((v) => v > z + 0.001) ?? 4 : [...ZOOMS].reverse().find((v) => v < z - 0.001) ?? 0.1;
    this.setZoom(next);
  }

  // ------------------------------------------------------------------ status / style bar

  syncStatus() {
    const n = this.doc?.slides().length ?? 0;
    $('sb-page').textContent = `슬라이드 ${n ? this.current + 1 : 0} / ${n}`;
    const s = this.doc?.slides()[this.current];
    $('sb-layout').textContent = s ? this.doc!.layoutName(s.path) : '';
    $('sb-mode').textContent = this.editor.insertMode ? '그리기' : this.editor.editing ? '텍스트 편집' : this.editor.selection.length ? `개체 ${this.editor.selection.length}개 선택` : '선택';
    $('sb-zoom-val').textContent = Math.round(this.editor.zoom * 100) + '%';
  }

  syncStyleBar() {
    this.syncStatus();
    const st = this.editor.caretStyle();
    const fn = $<HTMLInputElement>('font-name'), fs = $<HTMLInputElement>('font-size');
    if (document.activeElement !== fn) fn.value = st?.font ?? '';
    if (document.activeElement !== fs) fs.value = st ? String(st.size) : '';
    $('btn-bold').classList.toggle('active', !!st?.b);
    $('btn-italic').classList.toggle('active', !!st?.i);
    $('btn-underline').classList.toggle('active', !!st?.u);
    $('btn-strike').classList.toggle('active', !!st?.strike);
    for (const [id, a] of [['btn-align-left', 'l'], ['btn-align-center', 'ctr'], ['btn-align-right', 'r'], ['btn-align-justify', 'just']]) {
      $(id).classList.toggle('active', st?.algn === a);
    }
  }

  syncRulers() {
    if (!this.doc || !this.showRuler) return;
    const s = this.editor.selectedShapes()[0];
    drawRulers($<HTMLCanvasElement>('h-ruler'), $<HTMLCanvasElement>('v-ruler'), $('slide-host'), this.editor.zoom, this.doc.slideSize, s ? s.xfrm : null);
  }

  // ------------------------------------------------------------------ helpers for commands

  private get slidePath() { return this.doc?.slides()[this.current]?.path ?? ''; }
  private hasDoc = () => !!this.doc;
  private hasSlides = () => !!this.doc && this.doc.slides().length > 0;

  private selectedSlideList(): number[] {
    const n = this.doc?.slides().length ?? 0;
    const idx = [...this.slideSel].filter((i) => i < n).sort((a, b) => a - b);
    return idx.length ? idx : n ? [this.current] : [];
  }

  private selectSlideRange(start: number, count: number) {
    this.slideSel = new Set(Array.from({ length: count }, (_, k) => start + k));
    this.current = start;
    this.editor.selection = [];
    this.onDocChange();
    this.thumbs.markCurrent();
  }

  /** New slide at `at`. After a title slide PowerPoint continues with "Title and Content". */
  private insertNewSlide(at: number) {
    const doc = this.doc!;
    const lays = doc.layouts();
    const ref = doc.slides()[Math.max(0, at - 1)]?.path;
    const refLay = ref ? doc.pkg.rels(ref).find((r) => r.type.endsWith('/slideLayout'))?.target : undefined;
    const next = !refLay ? lays[1] ?? lays[0] : refLay === lays[0]?.path ? lays[1] ?? lays[0] : lays.find((l) => l.path === refLay) ?? lays[0];
    if (!next) return;
    const i = doc.addSlide(next.path, at);
    this.selectSlideRange(i, 1);
  }

  private deleteSlides(idx: number[]) {
    if (this.editor.editing) this.editor.commitEdit();
    this.doc!.deleteSlides(idx);
    const n = this.doc!.slides().length;
    this.selectSlideRange(Math.max(0, Math.min(idx[0], n - 1)), n ? 1 : 0);
  }

  private slideContextMenu(i: number, x: number, y: number) {
    if (!this.doc) return;
    this.thumbs.focusList();
    const en = (id: string) => !this.cmd.enabled(id);
    const run = (id: string) => () => this.cmd.run(id);
    const many = i >= 0 && this.slideSel.size > 1 ? ` (${this.slideSel.size}장)` : '';
    if (i < 0) {
      contextMenu(x, y, [
        { label: '새 슬라이드', shortcut: 'Ctrl+M', action: () => this.insertNewSlide(this.doc!.slides().length) },
        { label: '새 슬라이드 (레이아웃 선택)…', action: run('slide:new-with-layout') },
        { sep: true },
        { label: '붙여넣기', shortcut: 'Ctrl+V', disabled: en('slide:paste'), action: () => { this.current = Math.max(0, this.doc!.slides().length - 1); this.cmd.run('slide:paste'); } },
      ]);
      return;
    }
    contextMenu(x, y, [
      { label: '새 슬라이드', shortcut: 'Ctrl+M', action: () => this.insertNewSlide(i + 1) },
      { label: '새 슬라이드 (레이아웃 선택)…', action: run('slide:new-with-layout') },
      { label: '슬라이드 복제' + many, shortcut: 'Ctrl+D', action: run('slide:duplicate') },
      { sep: true },
      { label: '오려두기' + many, shortcut: 'Ctrl+X', action: run('slide:cut') },
      { label: '복사' + many, shortcut: 'Ctrl+C', action: run('slide:copy') },
      { label: '붙여넣기', shortcut: 'Ctrl+V', disabled: en('slide:paste'), action: run('slide:paste') },
      { sep: true },
      { label: '슬라이드 삭제' + many, shortcut: 'Delete', action: run('slide:delete') },
      { sep: true },
      { label: '레이아웃 변경…', action: run('slide:layout') },
      { label: '앞으로 이동', disabled: en('slide:move-up'), action: run('slide:move-up') },
      { label: '뒤로 이동', disabled: en('slide:move-down'), action: run('slide:move-down') },
    ]);
  }
  private hasSel = () => !!this.doc && this.editor.selection.length > 0;
  private editableSel() { return this.editor.selectedShapes().filter((s) => !s.locked).map((s) => s.id); }

  private toggleChar(key: 'b' | 'i' | 'u' | 'strike') {
    const st = this.editor.caretStyle();
    this.editor.applyCharProps({ [key]: !(st?.[key] ?? false) } as RunProps);
  }

  private defaultRect(wCm: number, hCm: number): Xfrm {
    const { cx, cy } = this.doc!.slideSize;
    const w = wCm * 360000, h = hCm * 360000;
    return { x: (cx - w) / 2, y: (cy - h) / 2, w, h, rot: 0, flipH: false, flipV: false };
  }

  private async insertImageFile(file: File | Blob, name = '그림') {
    const doc = this.doc!;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(file);
    const img = new Image();
    await new Promise<void>((res) => { img.onload = () => res(); img.onerror = () => res(); img.src = url; });
    URL.revokeObjectURL(url);
    const { cx, cy } = doc.slideSize;
    let w = (img.naturalWidth || 400) * 9525, h = (img.naturalHeight || 300) * 9525; // 96 dpi
    const k = Math.min(1, (cx * 0.8) / w, (cy * 0.8) / h);
    w *= k; h *= k;
    const id = doc.addPicture(this.slidePath, bytes, file.type || 'image/png', { x: (cx - w) / 2, y: (cy - h) / 2, w, h, rot: 0, flipH: false, flipV: false }, name);
    await this.images.ensureForSlide(this.slidePath);
    this.renderCurrent();
    this.thumbs.update([this.slidePath]);
    this.editor.select([id]);
  }

  private copyShapes(cut: boolean) {
    const ids = this.editableSel();
    if (!ids.length) return;
    const items = this.doc!.copyShapes(this.slidePath, ids);
    const text = this.editor.selectedShapes().map((s) => (s.text?.rendered.content.innerText ?? '').trim()).filter(Boolean).join('\n') || ' ';
    this.clipboard = { items, text, pasteCount: 0 };
    navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) this.doc!.deleteShapes(this.slidePath, ids);
  }

  private pasteShapes() {
    if (!this.clipboard) return;
    this.clipboard.pasteCount++;
    const ids = this.doc!.pasteShapes(this.slidePath, this.clipboard.items, this.clipboard.pasteCount * 182880);
    this.editor.select(ids);
  }

  private alignShapes(how: string) {
    const doc = this.doc!;
    const sel = this.editor.selectedShapes().filter((s) => !s.locked);
    if (!sel.length) return;
    const { cx, cy } = doc.slideSize;
    // One shape aligns to the slide; several align to their common bounds.
    const b = sel.length === 1 ? { x0: 0, y0: 0, x1: cx, y1: cy } : {
      x0: Math.min(...sel.map((s) => s.xfrm.x)), y0: Math.min(...sel.map((s) => s.xfrm.y)),
      x1: Math.max(...sel.map((s) => s.xfrm.x + s.xfrm.w)), y1: Math.max(...sel.map((s) => s.xfrm.y + s.xfrm.h)),
    };
    doc.setXfrms(this.slidePath, sel.map((s) => {
      const x = { ...s.xfrm };
      if (how === 'align-left') x.x = b.x0;
      if (how === 'align-hcenter') x.x = (b.x0 + b.x1 - x.w) / 2;
      if (how === 'align-right') x.x = b.x1 - x.w;
      if (how === 'align-top') x.y = b.y0;
      if (how === 'align-vcenter') x.y = (b.y0 + b.y1 - x.h) / 2;
      if (how === 'align-bottom') x.y = b.y1 - x.h;
      return { id: s.id, x };
    }), '맞춤');
  }

  private transformShapes(f: (x: Xfrm) => Xfrm, label: string) {
    const sel = this.editor.selectedShapes().filter((s) => !s.locked);
    if (sel.length) this.doc!.setXfrms(this.slidePath, sel.map((s) => ({ id: s.id, x: f({ ...s.xfrm }) })), label);
  }

  private nudge(dx: number, dy: number) {
    this.transformShapes((x) => ({ ...x, x: x.x + dx, y: x.y + dy }), '개체 이동');
  }

  private async save(saveAs: boolean) {
    if (!this.doc) return;
    if (this.editor.editing) this.editor.commitEdit();
    const blob = await this.doc.save();
    const name = this.fileName.toLowerCase().endsWith('.pptx') ? this.fileName : this.fileName.replace(/\.[^.]*$/, '') + '.pptx';
    const res = await saveFile(blob, name, this.fileHandle, saveAs, this.filePath);
    if (!res) { this.doc.dirty = true; return; }
    this.fileName = res.name;
    this.fileHandle = res.handle;
    this.filePath = res.path;
    this.updateTitle();
    this.message(`저장됨: ${res.name}`);
    toast('저장했습니다');
  }

  private findDialog() {
    const body = html(`<div class="form" style="grid-template-columns:auto 1fr">
      <label>찾을 내용</label><input type="text" name="f">
      <label>바꿀 내용</label><input type="text" name="r">
      <label></label><label style="color:inherit"><input type="checkbox" name="mc"> 대/소문자 구분</label>
      <label></label><span id="find-result" style="font-size:12px;color:#5c6370"></span></div>`);
    const q = (n: string) => body.querySelector(`[name="${n}"]`) as HTMLInputElement;
    const sel = window.getSelection()?.toString();
    if (sel) q('f').value = sel;
    let pos = -1;
    const result = body.querySelector('#find-result')!;
    dialog('찾기 / 바꾸기', body, [
      {
        label: '다음 찾기', primary: true, action: () => {
          const hits = this.doc!.findAll(q('f').value, q('mc').checked);
          if (!hits.length) { result.textContent = '찾는 내용이 없습니다.'; return false; }
          pos = (pos + 1) % hits.length;
          const h = hits[pos];
          this.gotoSlide(h.slide);
          this.editor.select([h.shape]);
          this.editor.shape(h.shape)?.el.classList.add('find-hit');
          result.textContent = `${hits.length}개 중 ${pos + 1}번째 (슬라이드 ${h.slide + 1})`;
          return false;
        },
      },
      {
        label: '모두 바꾸기', action: () => {
          const n = this.doc!.replaceAll(q('f').value, q('r').value, q('mc').checked);
          result.textContent = n ? `${n}개를 바꿨습니다.` : '찾는 내용이 없습니다.';
          return false;
        },
      },
      { label: '닫기' },
    ]);
  }

  // ------------------------------------------------------------------ commands

  private registerCommands() {
    const c = this.cmd;
    const d = () => this.doc!;
    // file
    c.register('file:new', () => this.confirmDiscard() && this.newDocument(), 'ctrl+n');
    c.register('file:open', async () => {
      if (!this.confirmDiscard()) return;
      const f = await pickOpenFile($<HTMLInputElement>('file-input'));
      if (f) await this.open(f.bytes, f.name, f.handle, f.path);
    }, 'ctrl+o');
    c.register('file:save', { run: () => this.save(false), enabled: this.hasDoc }, 'ctrl+s');
    c.register('file:save-as', { run: () => this.save(true), enabled: this.hasDoc }, 'ctrl+shift+s');
    c.register('file:print', { run: () => printSlides(d(), this.images), enabled: this.hasDoc }, 'ctrl+p');
    c.register('file:info', { run: () => this.infoDialog(), enabled: this.hasDoc });
    c.register('file:about', () => { dialog('제품 정보', html(`<div style="line-height:1.7"><b>Opptx</b> 0.1.0<br>PPTX 프레젠테이션 편집기<br><span style="color:#5c6370;font-size:12px">HOP(HWP 편집기)의 구조 — 문서 엔진 · 렌더러 · 명령 기반 UI — 를 참고해 만든 PowerPoint 편집기입니다.</span></div>`), [{ label: '확인', primary: true }]); });
    // edit
    c.register('edit:undo', { run: () => { if (this.editor.editing) { document.execCommand('undo'); return; } d().undo(); }, enabled: () => !!this.doc && (this.editor.editing != null || d().canUndo()) }, 'ctrl+z');
    c.register('edit:redo', { run: () => { if (this.editor.editing) { document.execCommand('redo'); return; } d().redo(); }, enabled: () => !!this.doc && (this.editor.editing != null || d().canRedo()) }, 'ctrl+shift+z', 'ctrl+y');
    c.register('edit:cut', { run: () => this.editor.editing ? document.execCommand('cut') : this.copyShapes(true), enabled: this.hasSel }, 'ctrl+x');
    c.register('edit:copy', { run: () => this.editor.editing ? document.execCommand('copy') : this.copyShapes(false), enabled: this.hasSel }, 'ctrl+c');
    c.register('edit:paste', { run: () => this.pasteShapes(), enabled: () => !!this.clipboard && !!this.doc && !this.editor.editing }, 'ctrl+v');
    c.register('edit:duplicate', { run: () => { this.copyShapes(false); this.pasteShapes(); }, enabled: this.hasSel }, 'ctrl+d');
    c.register('edit:delete', { run: () => d().deleteShapes(this.slidePath, this.editableSel()), enabled: this.hasSel }, 'delete', 'backspace');
    c.register('edit:select-all', { run: () => this.editor.select(this.editor.rendered?.shapes.filter((s) => !s.locked).map((s) => s.id) ?? []), enabled: this.hasDoc }, 'ctrl+a');
    c.register('edit:find', { run: () => this.findDialog(), enabled: this.hasDoc }, 'ctrl+f', 'ctrl+h');
    // view
    c.register('view:zoom-in', { run: () => this.stepZoom(1), enabled: this.hasDoc }, 'ctrl+=');
    c.register('view:zoom-out', { run: () => this.stepZoom(-1), enabled: this.hasDoc }, 'ctrl+-');
    c.register('view:zoom-fit', { run: () => this.zoomFit(), enabled: this.hasDoc }, 'ctrl+0');
    c.register('view:zoom-100', { run: () => this.setZoom(1), enabled: this.hasDoc });
    c.register('view:toggle-ruler', { run: () => { this.showRuler = !this.showRuler; $('editor-area').classList.toggle('no-ruler', !this.showRuler); this.syncRulers(); }, checked: () => this.showRuler });
    c.register('view:toggle-panel', { run: () => { this.showPanel = !this.showPanel; $('slide-panel').classList.toggle('hidden', !this.showPanel); if (this.fitZoom) this.zoomFit(); }, checked: () => this.showPanel });
    c.register('view:toggle-notes', { run: () => { this.showNotes = !this.showNotes; $('notes-pane').hidden = !this.showNotes; if (this.fitZoom) this.zoomFit(); }, checked: () => this.showNotes });
    c.register('view:toggle-grid', { run: () => { this.editor.showGrid = !this.editor.showGrid; this.renderCurrent(); }, checked: () => this.editor.showGrid });
    c.register('view:show-start', { run: () => startSlideShow(d(), this.images, 0, (i) => this.gotoSlide(i)), enabled: this.hasDoc }, 'f5');
    c.register('view:show-current', { run: () => startSlideShow(d(), this.images, this.current, (i) => this.gotoSlide(i)), enabled: this.hasDoc }, 'shift+f5');
    // insert
    c.register('insert:textbox', { run: () => this.editor.startInsert({ kind: 'textbox' }), enabled: this.hasDoc }, 'ctrl+shift+t');
    c.register('insert:shape', { run: () => shapePopover($('tb-shape'), (prst, label) => this.editor.startInsert({ kind: 'shape', prst, label })), enabled: this.hasDoc });
    c.register('insert:image', {
      run: async () => { const f = await pickImage($<HTMLInputElement>('image-input')); if (f) await this.insertImageFile(f, f.name); },
      enabled: this.hasDoc,
    }, 'ctrl+shift+i');
    c.register('insert:table', {
      run: () => tablePopover($('tb-table'), (r, col) => { const id = d().addTable(this.slidePath, r, col, this.defaultRect(Math.min(24, col * 3.5), r * 1)); this.editor.select([id]); }),
      enabled: this.hasDoc,
    });
    // format
    c.register('format:bold', { run: () => this.toggleChar('b'), enabled: this.hasSel }, 'ctrl+b');
    c.register('format:italic', { run: () => this.toggleChar('i'), enabled: this.hasSel }, 'ctrl+i');
    c.register('format:underline', { run: () => this.toggleChar('u'), enabled: this.hasSel }, 'ctrl+u');
    c.register('format:strike', { run: () => this.toggleChar('strike'), enabled: this.hasSel }, 'ctrl+shift+x');
    c.register('format:align-left', { run: () => this.editor.align('l'), enabled: this.hasSel }, 'ctrl+l');
    c.register('format:align-center', { run: () => this.editor.align('ctr'), enabled: this.hasSel }, 'ctrl+e');
    c.register('format:align-right', { run: () => this.editor.align('r'), enabled: this.hasSel }, 'ctrl+r');
    c.register('format:align-justify', { run: () => this.editor.align('just'), enabled: this.hasSel }, 'ctrl+j');
    c.register('format:bullet', { run: () => this.editor.paraOp({ bullet: 'char' }), enabled: this.hasSel });
    c.register('format:numbering', { run: () => this.editor.paraOp({ bullet: 'num' }), enabled: this.hasSel });
    c.register('format:level-increase', { run: () => this.editor.paraOp({ lvlDelta: 1 }), enabled: this.hasSel });
    c.register('format:level-decrease', { run: () => this.editor.paraOp({ lvlDelta: -1 }), enabled: this.hasSel });
    c.register('format:fill', { run: () => colorPopover($('tb-fill'), { allowNone: '채우기 없음' }, (hex) => d().setFill(this.slidePath, this.editableSel(), hex)), enabled: this.hasSel });
    c.register('format:line', { run: () => colorPopover($('tb-line'), { allowNone: '선 없음' }, (hex) => d().setLine(this.slidePath, this.editableSel(), hex)), enabled: this.hasSel });
    c.register('format:object-properties', {
      run: () => {
        const s = this.editor.selectedShapes()[0];
        if (!s) return;
        objectPropsDialog(s.xfrm, (r) => {
          const ids = this.editableSel();
          d().transact('개체 속성', [this.slidePath], () => {
            d().setXfrms(this.slidePath, [{ id: s.id, x: r.x }]);
            if (r.fill !== undefined) d().setFill(this.slidePath, ids, r.fill);
            if (r.line !== undefined || r.lineW !== undefined) d().setLine(this.slidePath, ids, r.line === undefined ? '000000' : r.line, r.lineW);
          });
        });
      },
      enabled: this.hasSel,
    }, 'ctrl+shift+p');
    // slide
    c.register('slide:new', { run: () => this.insertNewSlide(d().slides().length ? this.current + 1 : 0), enabled: this.hasDoc }, 'ctrl+m');
    c.register('slide:duplicate', {
      run: () => {
        const idx = this.selectedSlideList();
        const at = d().duplicateSlides(idx);
        this.selectSlideRange(at, idx.length);
        this.message(`슬라이드 ${idx.length}장을 복제했습니다`);
      },
      enabled: this.hasSlides,
    });
    c.register('slide:copy', {
      run: () => { const idx = this.selectedSlideList(); this.slideClipboard = d().copySlides(idx); this.message(`슬라이드 ${idx.length}장을 복사했습니다`); },
      enabled: this.hasSlides,
    });
    c.register('slide:cut', {
      run: () => {
        const idx = this.selectedSlideList();
        this.slideClipboard = d().copySlides(idx);
        this.deleteSlides(idx);
        this.message(`슬라이드 ${idx.length}장을 오려두었습니다`);
      },
      enabled: this.hasSlides,
    });
    c.register('slide:paste', {
      run: () => {
        const clip = this.slideClipboard!;
        const at = d().pasteSlides(clip, d().slides().length ? this.current + 1 : 0);
        this.selectSlideRange(at, clip.length);
        this.message(`슬라이드 ${clip.length}장을 붙여넣었습니다`);
      },
      enabled: () => !!this.doc && !!this.slideClipboard?.length,
    });
    c.register('slide:delete', {
      run: () => {
        const idx = this.selectedSlideList();
        this.deleteSlides(idx);
        this.message(`슬라이드 ${idx.length}장을 삭제했습니다 (Ctrl+Z로 되돌리기)`);
      },
      enabled: this.hasSlides,
    });
    c.register('slide:move-up', { run: () => this.moveSlides(this.current, this.current - 1), enabled: () => !!this.doc && this.current > 0 });
    c.register('slide:move-down', { run: () => this.moveSlides(this.current, this.current + 1), enabled: () => !!this.doc && this.current < d().slides().length - 1 });
    c.register('slide:layout', {
      run: () => layoutDialog(d(), '레이아웃', (p) => {
        if (!d().slides().length) { d().addSlide(p, 0); this.gotoSlide(0); return; }
        d().setSlideLayout(this.slidePath, p);
      }),
      enabled: this.hasDoc,
    });
    c.register('slide:new-with-layout', { run: () => layoutDialog(d(), '새 슬라이드', (p) => this.gotoSlide(d().addSlide(p, d().slides().length ? this.current + 1 : 0))), enabled: this.hasDoc });
    c.register('slide:background', { run: () => colorPopover($('tb-fill'), { allowNone: '기본 배경' }, (hex) => d().setSlideBackground(this.slidePath, hex)), enabled: () => !!this.slidePath });
    // arrange
    for (const how of ['front', 'back', 'forward', 'backward'] as const) {
      c.register(`arrange:${how}`, { run: () => d().reorder(this.slidePath, this.editableSel(), how), enabled: this.hasSel },
        ...({ front: ['ctrl+shift+]'], back: ['ctrl+shift+['], forward: ['ctrl+]'], backward: ['ctrl+['] }[how]));
    }
    for (const how of ['align-left', 'align-hcenter', 'align-right', 'align-top', 'align-vcenter', 'align-bottom']) {
      c.register(`arrange:${how}`, { run: () => this.alignShapes(how), enabled: this.hasSel });
    }
    c.register('arrange:flip-h', { run: () => this.transformShapes((x) => ({ ...x, flipH: !x.flipH }), '좌우 대칭'), enabled: this.hasSel });
    c.register('arrange:flip-v', { run: () => this.transformShapes((x) => ({ ...x, flipV: !x.flipV }), '상하 대칭'), enabled: this.hasSel });
    c.register('arrange:rotate-right', { run: () => this.transformShapes((x) => ({ ...x, rot: (x.rot + 90) % 360 }), '회전'), enabled: this.hasSel });
    // tools
    c.register('tool:shortcuts', () => {
      const labels = new Map<string, string>();
      document.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((m) => labels.set(m.dataset.cmd!, m.querySelector('.md-label')!.textContent!));
      const rows = c.shortcutList().filter(([, id]) => labels.has(id)).map(([k, id]) => `<tr><td>${labels.get(id)}</td><td><kbd>${k.replace(/\+/g, '</kbd> + <kbd>')}</kbd></td></tr>`).join('');
      dialog('단축키 목록', html(`<table class="kbd-table">${rows}<tr><td>개체 미세 이동</td><td><kbd>←↑→↓</kbd> (Shift: 크게)</td></tr><tr><td>텍스트 편집 시작 / 끝</td><td><kbd>Enter</kbd> / <kbd>Esc</kbd></td></tr></table>`), [{ label: '닫기', primary: true }]);
    }, 'f1');
  }

  private infoDialog() {
    const doc = this.doc!;
    const { cx, cy } = doc.slideSize;
    const core = doc.pkg.getXml('docProps/core.xml');
    const get = (n: string) => core?.getElementsByTagNameNS('*', n)[0]?.textContent ?? '';
    dialog('문서 정보', html(`<table class="kbd-table">
      <tr><td>파일</td><td>${this.fileName}</td></tr>
      <tr><td>슬라이드</td><td>${doc.slides().length}장</td></tr>
      <tr><td>크기</td><td>${(cx / 360000).toFixed(2)} × ${(cy / 360000).toFixed(2)} cm</td></tr>
      <tr><td>레이아웃</td><td>${doc.layouts().length}개</td></tr>
      <tr><td>제목</td><td>${get('title')}</td></tr>
      <tr><td>만든 이</td><td>${get('creator')}</td></tr>
      <tr><td>수정</td><td>${get('modified')}</td></tr></table>`), [{ label: '확인', primary: true }]);
  }

  private confirmDiscard() {
    return !this.doc?.dirty || confirm('저장하지 않은 변경 내용이 있습니다. 계속할까요?');
  }

  // ------------------------------------------------------------------ DOM bindings

  private bindUi() {
    installIcons();
    $('font-list').innerHTML = FONTS.map((f) => `<option value="${f}">`).join('');

    // Menus: click to open, hover to switch (like HOP's menu bar).
    const menus = document.querySelectorAll<HTMLElement>('.menu-item');
    const closeMenus = () => menus.forEach((m) => m.classList.remove('open'));
    const refreshMenu = (m: HTMLElement) => m.querySelectorAll<HTMLElement>('.md-item[data-cmd]').forEach((it) => {
      it.classList.toggle('disabled', !this.cmd.enabled(it.dataset.cmd!));
      it.classList.toggle('checked', this.cmd.checked(it.dataset.cmd!));
    });
    menus.forEach((m) => {
      m.querySelector('.menu-title')!.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const was = m.classList.contains('open');
        closeMenus();
        if (!was) { refreshMenu(m); m.classList.add('open'); }
      });
      m.addEventListener('mouseenter', () => {
        if ([...menus].some((x) => x.classList.contains('open')) && !m.classList.contains('open')) { closeMenus(); refreshMenu(m); m.classList.add('open'); }
      });
    });
    document.addEventListener('mousedown', (e) => { if (!(e.target as HTMLElement).closest('.menu-item')) closeMenus(); });

    // Any [data-cmd] element runs its command. mousedown is suppressed to keep the text caret.
    document.querySelectorAll<HTMLElement>('[data-cmd]').forEach((el) => {
      el.addEventListener('mousedown', (e) => e.preventDefault());
      el.addEventListener('click', () => { closeMenus(); this.cmd.run(el.dataset.cmd!); });
    });
    $('slide-panel-bar').addEventListener('mousedown', () => this.thumbs.focusList());
    $('slide-panel').addEventListener('dblclick', (e) => { if (!(e.target as HTMLElement).closest('.thumb')) this.cmd.run('slide:new'); });

    // Style bar inputs.
    const fn = $<HTMLInputElement>('font-name'), fs = $<HTMLInputElement>('font-size');
    for (const inp of [fn, fs]) {
      inp.addEventListener('mousedown', () => this.editor.rememberSelection());
      inp.addEventListener('focus', () => inp.select());
    }
    fn.addEventListener('change', () => { if (fn.value.trim()) this.editor.applyCharProps({ font: fn.value.trim() }); });
    fn.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); fn.blur(); } });
    const setSize = (v: number) => { if (v > 0 && v <= 4000) this.editor.applyCharProps({ sz: v }); };
    fs.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); setSize(parseFloat(fs.value)); fs.blur(); } });
    const step = (dir: number) => {
      const cur = this.editor.caretStyle()?.size ?? 18;
      const sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
      const next = dir > 0 ? sizes.find((s) => s > cur) ?? cur + 8 : [...sizes].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
      setSize(next);
    };
    for (const [id, dir] of [['btn-size-up', 1], ['btn-size-down', -1]] as const) {
      $(id).addEventListener('mousedown', (e) => e.preventDefault());
      $(id).addEventListener('click', () => step(dir));
    }
    const colorBtn = $('btn-text-color');
    colorBtn.addEventListener('mousedown', (e) => { e.preventDefault(); this.editor.rememberSelection(); });
    colorBtn.addEventListener('click', (e) => {
      const r = colorBtn.getBoundingClientRect();
      // Left part applies the last color, the rest opens the palette.
      if (e.clientX - r.left < r.width * 0.6 && e.detail) { this.editor.applyCharProps({ color: this.lastTextColor }); return; }
      colorPopover(colorBtn, {}, (hex) => {
        if (!hex) return;
        this.lastTextColor = hex;
        $('color-bar').style.background = '#' + hex;
        this.editor.applyCharProps({ color: hex });
      });
    });
    colorBtn.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      colorPopover(colorBtn, {}, (hex) => { if (hex) { this.lastTextColor = hex; $('color-bar').style.background = '#' + hex; this.editor.applyCharProps({ color: hex }); } });
    });

    // Keyboard.
    window.addEventListener('keydown', (e) => this.onKey(e));
    // Clipboard integration for images/text from other apps.
    document.addEventListener('paste', (e) => this.onPaste(e));

    // Drag & drop of .pptx files / images.
    const sc = $('scroll-container');
    sc.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('Files')) { e.preventDefault(); sc.classList.add('dragover'); } });
    sc.addEventListener('dragleave', () => sc.classList.remove('dragover'));
    sc.addEventListener('drop', async (e) => {
      e.preventDefault();
      sc.classList.remove('dragover');
      const f = e.dataTransfer?.files[0];
      if (!f) return;
      if (/\.pptx$/i.test(f.name)) { if (this.confirmDiscard()) await this.open(new Uint8Array(await f.arrayBuffer()), f.name); }
      else if (f.type.startsWith('image/') && this.doc) await this.insertImageFile(f, f.name);
    });
    sc.addEventListener('scroll', () => this.syncRulers());
    sc.addEventListener('wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey) || !this.doc) return;
      e.preventDefault();
      this.setZoom(this.editor.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    new ResizeObserver(() => { if (this.fitZoom) this.zoomFit(); else this.syncRulers(); }).observe(sc);
    window.addEventListener('beforeunload', (e) => { if (this.doc?.dirty) e.preventDefault(); });
  }

  private onKey(e: KeyboardEvent) {
    if (document.querySelector('.dlg-backdrop') || document.getElementById('slideshow')) return;
    const t = e.target as HTMLElement;
    const inInput = t.matches('input, textarea, select');
    if (e.key === 'Escape') {
      closePopover();
      closeContextMenu();
      if (this.editor.insertMode) { this.editor.startInsert(null); return; }
      if (!this.editor.editing && this.editor.selection.length) { this.editor.select([]); return; }
    }
    if (inInput) return;
    const editing = !!this.editor.editing;
    const slideListFocused = t.closest('#slide-panel') != null;

    if (slideListFocused && !editing) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); this.gotoSlide(this.current + (e.key === 'ArrowDown' ? 1 : -1), e.shiftKey ? 'range' : undefined); return; }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.cmd.run('slide:delete'); return; }
      const mod = e.ctrlKey || e.metaKey;
      const panelCmd = mod && !e.shiftKey && !e.altKey ? ({ KeyD: 'slide:duplicate', KeyC: 'slide:copy', KeyX: 'slide:cut', KeyV: 'slide:paste', KeyM: 'slide:new' } as Record<string, string>)[e.code] : undefined;
      if (panelCmd) { e.preventDefault(); this.cmd.run(panelCmd); return; }
      if (mod && e.code === 'KeyA') { e.preventDefault(); this.slideSel = new Set(this.doc!.slides().map((_, k) => k)); this.thumbs.markCurrent(); return; }
      if (e.key === 'Enter') { e.preventDefault(); this.cmd.run('slide:new'); return; }
    }

    const id = this.cmd.lookup(e);
    if (editing) {
      // Let the browser handle text-editing keys inside contenteditable.
      const passthrough = ['edit:undo', 'edit:redo', 'edit:cut', 'edit:copy', 'edit:paste', 'edit:delete', 'edit:select-all', 'edit:duplicate'];
      if (!id || passthrough.includes(id)) return;
      e.preventDefault();
      this.cmd.run(id);
      return;
    }
    if (id) {
      if (['edit:cut', 'edit:copy', 'edit:paste'].includes(id)) return; // handled by clipboard events
      e.preventDefault();
      this.cmd.run(id);
      return;
    }
    if (!this.doc) return;
    const sel = this.editor.selectedShapes();
    if (e.key.startsWith('Arrow') && sel.length) {
      e.preventDefault();
      const k = (e.shiftKey ? 10 : 1) * 12700 * (e.altKey ? 0.25 : 1);
      this.nudge(e.key === 'ArrowLeft' ? -k : e.key === 'ArrowRight' ? k : 0, e.key === 'ArrowUp' ? -k : e.key === 'ArrowDown' ? k : 0);
    } else if ((e.key === 'Enter' || e.key === 'F2') && sel.length === 1 && (sel[0].text || sel[0].cells.length)) {
      e.preventDefault();
      this.editor.enterEdit(sel[0].id);
    } else if (e.key === 'Tab' && this.editor.rendered) {
      e.preventDefault();
      const shapes = this.editor.rendered.shapes.filter((s) => !s.locked);
      if (!shapes.length) return;
      const i = shapes.findIndex((s) => s.id === sel[0]?.id);
      this.editor.select([shapes[(i + (e.shiftKey ? -1 : 1) + shapes.length) % shapes.length].id]);
    } else if (e.key === 'PageDown' || e.key === 'PageUp') {
      e.preventDefault();
      this.gotoSlide(this.current + (e.key === 'PageDown' ? 1 : -1));
    } else if (!sel.length && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault();
      this.gotoSlide(this.current + (e.key === 'ArrowDown' ? 1 : -1));
    }
  }

  private async onPaste(e: ClipboardEvent) {
    if (!this.doc || this.editor.editing || (e.target as HTMLElement).matches?.('input,textarea')) return;
    if (document.querySelector('.dlg-backdrop')) return;
    const items = Array.from(e.clipboardData?.items ?? []);
    const img = items.find((i) => i.type.startsWith('image/'));
    const text = e.clipboardData?.getData('text/plain') ?? '';
    e.preventDefault();
    if (this.clipboard && (text === this.clipboard.text || (!text && !img))) { this.pasteShapes(); return; }
    if (img) { const f = img.getAsFile(); if (f) await this.insertImageFile(f); return; }
    if (text.trim()) {
      const id = this.doc.addTextBox(this.slidePath, this.defaultRect(16, 1.2), text.trim().split(/\r?\n/)[0]);
      this.editor.select([id]);
    }
  }

  onCopyEvent(e: ClipboardEvent, cut: boolean) {
    if (!this.doc || this.editor.editing || !this.editor.selection.length) return;
    if ((e.target as HTMLElement).matches?.('input,textarea') || document.querySelector('.dlg-backdrop')) return;
    e.preventDefault();
    this.copyShapes(cut);
    e.clipboardData?.setData('text/plain', this.clipboard!.text);
  }
}

