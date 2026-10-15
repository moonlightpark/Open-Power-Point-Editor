// Horizontal/vertical rulers in centimetres, aligned with the slide (HOP: #h-ruler / #v-ruler canvases).
const EMU_PER_CM = 360000;

export function drawRulers(h: HTMLCanvasElement, v: HTMLCanvasElement, slideHost: HTMLElement, zoom: number, slideEmu: { cx: number; cy: number }, sel: { x: number; y: number; w: number; h: number } | null) {
  const dpr = window.devicePixelRatio || 1;
  const host = slideHost.getBoundingClientRect();
  const draw = (c: HTMLCanvasElement, horizontal: boolean) => {
    const r = c.getBoundingClientRect();
    if (!r.width || !r.height) return;
    c.width = r.width * dpr;
    c.height = r.height * dpr;
    const g = c.getContext('2d')!;
    g.scale(dpr, dpr);
    g.clearRect(0, 0, r.width, r.height);
    const origin = horizontal ? host.left - r.left : host.top - r.top;
    const len = ((horizontal ? slideEmu.cx : slideEmu.cy) / 12700) * zoom;
    const thick = horizontal ? r.height : r.width;
    g.fillStyle = '#fff';
    if (horizontal) g.fillRect(origin, 3, len, thick - 6); else g.fillRect(3, origin, thick - 6, len);
    if (sel) {
      g.fillStyle = 'rgba(43,111,214,.18)';
      const a = origin + ((horizontal ? sel.x : sel.y) / 12700) * zoom;
      const b = ((horizontal ? sel.w : sel.h) / 12700) * zoom;
      if (horizontal) g.fillRect(a, 3, b, thick - 6); else g.fillRect(3, a, thick - 6, b);
    }
    const cm = (EMU_PER_CM / 12700) * zoom;
    const total = (horizontal ? slideEmu.cx : slideEmu.cy) / EMU_PER_CM;
    g.strokeStyle = '#8b929c';
    g.fillStyle = '#5c6370';
    g.font = '9px -apple-system, sans-serif';
    g.textAlign = 'center';
    g.textBaseline = 'middle';
    const step = cm < 12 ? 2 : 1;
    g.beginPath();
    for (let i = 0; i <= total * 2; i++) {
      const pos = origin + (i / 2) * cm;
      const major = i % 2 === 0;
      const tl = major ? 6 : 3;
      if (horizontal) { g.moveTo(pos + 0.5, thick - tl); g.lineTo(pos + 0.5, thick); }
      else { g.moveTo(thick - tl, pos + 0.5); g.lineTo(thick, pos + 0.5); }
      if (major && (i / 2) % step === 0 && i > 0 && i / 2 < total) {
        if (horizontal) g.fillText(String(i / 2), pos, thick / 2 - 1);
        else { g.save(); g.translate(thick / 2 - 1, pos); g.rotate(-Math.PI / 2); g.fillText(String(i / 2), 0, 0); g.restore(); }
      }
    }
    g.stroke();
  };
  draw(h, true);
  draw(v, false);
}
