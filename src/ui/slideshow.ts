// Full-screen slide show.
import { PptxDocument } from '../engine/document';
import { ImageCache, renderSlide } from '../engine/render';

export function startSlideShow(doc: PptxDocument, images: ImageCache, start: number, onExit: (index: number) => void) {
  const slides = doc.slides();
  if (!slides.length) return;
  let i = Math.min(Math.max(0, start), slides.length - 1);
  const root = document.createElement('div');
  root.id = 'slideshow';
  root.tabIndex = 0;
  const counter = document.createElement('div');
  counter.className = 'show-counter';
  document.body.appendChild(root);

  const show = async () => {
    await images.ensureForSlide(slides[i].path);
    const r = renderSlide(doc, slides[i].path, images);
    const k = Math.min(window.innerWidth / r.width, window.innerHeight / r.height);
    r.el.style.transform = `scale(${k})`;
    r.el.style.left = (window.innerWidth - r.width * k) / 2 + 'px';
    r.el.style.top = (window.innerHeight - r.height * k) / 2 + 'px';
    counter.textContent = `${i + 1} / ${slides.length}`;
    root.replaceChildren(r.el, counter);
  };
  const exit = () => {
    window.removeEventListener('keydown', onKey, true);
    window.removeEventListener('resize', show);
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    root.remove();
    onExit(i);
  };
  const go = (d: number) => {
    const n = i + d;
    if (n >= slides.length) { exit(); return; }
    if (n < 0) return;
    i = n;
    show();
  };
  const onKey = (e: KeyboardEvent) => {
    e.preventDefault();
    e.stopPropagation();
    switch (e.key) {
      case 'Escape': exit(); break;
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': go(1); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': go(-1); break;
      case 'Home': i = 0; show(); break;
      case 'End': i = slides.length - 1; show(); break;
    }
  };
  root.addEventListener('click', () => go(1));
  root.addEventListener('contextmenu', (e) => { e.preventDefault(); go(-1); });
  window.addEventListener('keydown', onKey, true);
  window.addEventListener('resize', show);
  root.requestFullscreen?.().catch(() => {});
  document.addEventListener('fullscreenchange', function f() {
    if (!document.fullscreenElement && root.isConnected) { document.removeEventListener('fullscreenchange', f); exit(); }
  });
  show();
}

export async function printSlides(doc: PptxDocument, images: ImageCache) {
  const root = document.getElementById('print-root')!;
  root.replaceChildren();
  const { cx, cy } = doc.slideSize;
  const w = cx / 12700, h = cy / 12700;
  const style = document.createElement('style');
  style.textContent = `@page { size: ${w}pt ${h}pt; margin: 0 }`;
  root.appendChild(style);
  for (const s of doc.slides()) {
    await images.ensureForSlide(s.path);
    const page = document.createElement('div');
    page.className = 'print-page';
    page.style.width = w + 'pt';
    page.style.height = h + 'pt';
    const r = renderSlide(doc, s.path, images);
    r.el.style.transform = 'scale(1.3333333)';
    page.appendChild(r.el);
    root.appendChild(page);
  }
  // Give images a moment to decode before the print dialog snapshots the page.
  await Promise.all(Array.from(root.querySelectorAll('img')).map((img) => img.decode().catch(() => {})));
  window.print();
  root.replaceChildren();
}
