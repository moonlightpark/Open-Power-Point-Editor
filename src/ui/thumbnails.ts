// Slide list panel: thumbnails, selection, drag-to-reorder.
import { PptxDocument } from '../engine/document';
import { ImageCache, renderSlide } from '../engine/render';

export interface ThumbHost {
  doc(): PptxDocument | null;
  images(): ImageCache;
  current(): number;
  goto(i: number, extend?: 'toggle' | 'range'): void;
  selectedSlides(): Set<number>;
  move(from: number, to: number): void;
  /** Right-click menu for a slide (index) or the empty panel area (-1). */
  contextMenu(index: number, x: number, y: number): void;
  /** Hover buttons on a thumbnail. */
  action(cmd: 'duplicate' | 'delete' | 'new-after', index: number): void;
}

const THUMB_W = 160;

export class Thumbnails {
  private cache = new Map<string, HTMLElement>(); // slide path -> frame
  constructor(private host: ThumbHost, private list: HTMLElement) {
    list.addEventListener('contextmenu', (e) => {
      if ((e.target as HTMLElement).closest('.thumb')) return;
      e.preventDefault();
      this.host.contextMenu(-1, e.clientX, e.clientY);
    });
  }

  /** Re-render thumbnails. `paths` limits which slides are re-rendered (others reuse cached DOM). */
  update(paths?: string[]) {
    const doc = this.host.doc();
    if (!doc) { this.list.replaceChildren(); this.cache.clear(); return; }
    const slides = doc.slides();
    const { cx, cy } = doc.slideSize;
    const scale = THUMB_W / (cx / 12700);
    const h = (cy / 12700) * scale;
    const keep = new Set(slides.map((s) => s.path));
    for (const k of [...this.cache.keys()]) if (!keep.has(k)) this.cache.delete(k);
    const items: HTMLElement[] = [];
    const sel = this.host.selectedSlides();
    slides.forEach((s, i) => {
      let frame = this.cache.get(s.path);
      if (!frame || !paths || paths.includes(s.path)) {
        frame = document.createElement('div');
        frame.className = 'thumb-frame';
        frame.style.height = h + 'px';
        const r = renderSlide(doc, s.path, this.host.images());
        r.el.style.transform = `scale(${scale})`;
        frame.appendChild(r.el);
        this.cache.set(s.path, frame);
      }
      const item = document.createElement('div');
      item.className = 'thumb' + (i === this.host.current() ? ' current' : '') + (sel.has(i) ? ' selected' : '');
      item.dataset.index = String(i);
      item.draggable = true;
      const num = document.createElement('div');
      num.className = 'thumb-num';
      num.textContent = String(i + 1);
      item.append(num, frame);
      const tools = document.createElement('div');
      tools.className = 'thumb-tools';
      for (const [cmd, title, glyph] of [['duplicate', '슬라이드 복제', '⧉'], ['delete', '슬라이드 삭제', '✕']] as const) {
        const b = document.createElement('button');
        b.className = 'thumb-tool';
        b.title = title;
        b.textContent = glyph;
        b.addEventListener('mousedown', (e) => { e.stopPropagation(); e.preventDefault(); this.focusList(); });
        b.addEventListener('click', (e) => { e.stopPropagation(); this.host.action(cmd, i); });
        tools.appendChild(b);
      }
      frame.querySelector('.thumb-tools')?.remove();
      frame.appendChild(tools);
      item.addEventListener('mousedown', (e) => {
        // Right-click keeps an existing multi-selection so the menu acts on all of it.
        if (e.button === 2 && sel.has(i)) return;
        this.host.goto(i, e.shiftKey ? 'range' : e.metaKey || e.ctrlKey ? 'toggle' : undefined);
      });
      item.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); this.host.contextMenu(i, e.clientX, e.clientY); });
      item.addEventListener('dragstart', (e) => { e.dataTransfer!.setData('application/x-opptx-slide', String(i)); e.dataTransfer!.effectAllowed = 'move'; });
      item.addEventListener('dragover', (e) => {
        if (!e.dataTransfer!.types.includes('application/x-opptx-slide')) return;
        e.preventDefault();
        const r = item.getBoundingClientRect();
        const after = e.clientY > r.top + r.height / 2;
        this.clearDrop();
        item.classList.add(after ? 'drop-after' : 'drop-before');
      });
      item.addEventListener('dragleave', () => item.classList.remove('drop-after', 'drop-before'));
      item.addEventListener('drop', (e) => {
        e.preventDefault();
        const from = Number(e.dataTransfer!.getData('application/x-opptx-slide'));
        const after = item.classList.contains('drop-after');
        this.clearDrop();
        let to = i + (after ? 1 : 0);
        if (from < to) to--;
        if (from !== to) this.host.move(from, to);
      });
      items.push(item);
      const gap = document.createElement('div');
      gap.className = 'thumb-gap';
      gap.title = '여기에 새 슬라이드 추가';
      gap.innerHTML = '<span>+</span>';
      gap.addEventListener('mousedown', (e) => { e.preventDefault(); this.focusList(); });
      gap.addEventListener('click', () => this.host.action('new-after', i));
      items.push(gap);
    });
    this.list.replaceChildren(...items);
    this.list.querySelector('.thumb.current')?.scrollIntoView({ block: 'nearest' });
  }

  /** Keep keyboard shortcuts (Ctrl+C/V/D, Delete) routed to the slide list after panel clicks. */
  focusList() {
    this.list.focus({ preventScroll: true });
  }

  private clearDrop() {
    this.list.querySelectorAll('.drop-after,.drop-before').forEach((x) => x.classList.remove('drop-after', 'drop-before'));
  }

  markCurrent() {
    const sel = this.host.selectedSlides();
    this.list.querySelectorAll<HTMLElement>('.thumb').forEach((t) => {
      const i = Number(t.dataset.index);
      t.classList.toggle('current', i === this.host.current());
      t.classList.toggle('selected', sel.has(i));
    });
    this.list.querySelector('.thumb.current')?.scrollIntoView({ block: 'nearest' });
  }

  clear() { this.cache.clear(); }
}
