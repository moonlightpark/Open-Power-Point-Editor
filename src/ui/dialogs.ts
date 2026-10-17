// Popovers (color palette, shape picker, table grid) and modal dialogs.
import { PAGE_SIZES, PageSizeKey, PptxDocument, Xfrm, placeholder, readXfrm, shapeChildren, spTree, xfrmEl, findInherited } from '../engine/document';
import { INSERTABLE, presetPath } from '../engine/geometry';

let openPop: HTMLElement | null = null;

export function closePopover() {
  openPop?.remove();
  openPop = null;
}

function popover(anchor: HTMLElement, content: HTMLElement) {
  closePopover();
  const pop = document.createElement('div');
  pop.className = 'popover';
  pop.appendChild(content);
  pop.addEventListener('mousedown', (e) => { if (!(e.target as HTMLElement).matches('input')) e.preventDefault(); });
  document.body.appendChild(pop);
  const r = anchor.getBoundingClientRect();
  const pw = pop.offsetWidth;
  pop.style.left = Math.min(r.left, window.innerWidth - pw - 8) + 'px';
  pop.style.top = r.bottom + 4 + 'px';
  openPop = pop;
  setTimeout(() => {
    const off = (e: MouseEvent) => {
      if (!pop.contains(e.target as Node)) { closePopover(); document.removeEventListener('mousedown', off, true); }
    };
    document.addEventListener('mousedown', off, true);
  });
  return pop;
}

const PALETTE = [
  'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
  'F2F2F2', '7F7F7F', 'D0CECE', 'D6DCE4', 'D9E2F3', 'FBE5D5', 'EDEDED', 'FFF2CC', 'DEEBF6', 'E2EFD9',
  'D8D8D8', '595959', 'AEABAB', 'ADB9CA', 'B4C6E7', 'F7CBAC', 'DBDBDB', 'FEE599', 'BDD7EE', 'C5E0B3',
  'BFBFBF', '3F3F3F', '757070', '8496B0', '8EAADB', 'F4B183', 'C9C9C9', 'FFD965', '9CC3E5', 'A8D08D',
  'A5A5A5', '262626', '3A3838', '323F4F', '2F5496', 'C55A11', '7B7B7B', 'BF9000', '2E75B5', '538135',
  'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0',
];

export function colorPopover(anchor: HTMLElement, opts: { allowNone?: string }, pick: (hex: string | null) => void) {
  const wrap = document.createElement('div');
  const grid = document.createElement('div');
  grid.className = 'palette';
  for (const c of PALETTE) {
    const b = document.createElement('button');
    b.className = 'swatch';
    b.style.background = '#' + c;
    b.title = '#' + c;
    b.onclick = () => { closePopover(); pick(c); };
    grid.appendChild(b);
  }
  wrap.appendChild(grid);
  const row = document.createElement('div');
  row.className = 'pop-row';
  if (opts.allowNone) {
    const none = document.createElement('button');
    none.textContent = opts.allowNone;
    none.onclick = () => { closePopover(); pick(null); };
    row.appendChild(none);
  }
  const custom = document.createElement('input');
  custom.type = 'color';
  custom.title = '다른 색';
  custom.onchange = () => { closePopover(); pick(custom.value.slice(1).toUpperCase()); };
  const lbl = document.createElement('span');
  lbl.textContent = '다른 색:';
  row.append(lbl, custom);
  wrap.appendChild(row);
  popover(anchor, wrap);
}

export function shapePopover(anchor: HTMLElement, pick: (prst: string, label: string) => void) {
  const grid = document.createElement('div');
  grid.className = 'shape-grid';
  for (const s of INSERTABLE) {
    const b = document.createElement('button');
    b.className = 'shape-cell';
    b.title = s.label;
    const d = presetPath(s.id, 26, 26, null);
    b.innerHTML = `<svg width="26" height="26" viewBox="-1 -1 28 28"><path d="${d}" fill="${s.id === 'line' ? 'none' : '#cfe0fb'}" stroke="#2b6fd6" stroke-width="1.3"/></svg>`;
    b.onclick = () => { closePopover(); pick(s.id, s.label); };
    grid.appendChild(b);
  }
  popover(anchor, grid);
}

export function tablePopover(anchor: HTMLElement, pick: (rows: number, cols: number) => void) {
  const wrap = document.createElement('div');
  const grid = document.createElement('div');
  grid.className = 'table-grid';
  const label = document.createElement('div');
  label.className = 'table-label';
  label.textContent = '표 크기 선택';
  const cells: HTMLElement[] = [];
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const d = document.createElement('div');
    d.onmouseenter = () => {
      cells.forEach((x, i) => x.classList.toggle('on', Math.floor(i / 10) <= r && i % 10 <= c));
      label.textContent = `${r + 1} × ${c + 1} 표`;
    };
    d.onclick = () => { closePopover(); pick(r + 1, c + 1); };
    cells.push(d);
    grid.appendChild(d);
  }
  wrap.append(grid, label);
  popover(anchor, wrap);
}

// ------------------------------------------------------------------ modal dialogs

export function dialog(title: string, body: HTMLElement, buttons: { label: string; primary?: boolean; action?: () => boolean | void }[]): () => void {
  const back = document.createElement('div');
  back.className = 'dlg-backdrop';
  const dlg = document.createElement('div');
  dlg.className = 'dlg';
  const t = document.createElement('div');
  t.className = 'dlg-title';
  t.textContent = title;
  const b = document.createElement('div');
  b.className = 'dlg-body';
  b.appendChild(body);
  const f = document.createElement('div');
  f.className = 'dlg-foot';
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey, true); };
  for (const btn of buttons) {
    const e = document.createElement('button');
    e.textContent = btn.label;
    if (btn.primary) e.className = 'primary';
    e.onclick = () => { if (btn.action?.() !== false) close(); };
    f.appendChild(e);
  }
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    else if (e.key === 'Enter' && !(e.target as HTMLElement).matches('textarea')) {
      const primary = buttons.find((x) => x.primary);
      if (primary) { e.preventDefault(); e.stopPropagation(); if (primary.action?.() !== false) close(); }
    }
    e.stopPropagation();
  };
  document.addEventListener('keydown', onKey, true);
  dlg.append(t, b);
  if (buttons.length) dlg.appendChild(f);
  back.appendChild(dlg);
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
  document.body.appendChild(back);
  setTimeout(() => (b.querySelector('input,button') as HTMLElement | null)?.focus());
  return close;
}

export function html(s: string): HTMLElement {
  const d = document.createElement('div');
  d.innerHTML = s;
  return d;
}

// ---------------------------------------------------------------- page format

const DEFAULT_PAGE_KEY = 'opptx.defaultPage';

/** Page format for new documents (A4 unless the user picked another one in 옵션). */
export function defaultPageSize(): PageSizeKey {
  try {
    const v = localStorage.getItem(DEFAULT_PAGE_KEY);
    if (v && v in PAGE_SIZES) return v as PageSizeKey;
  } catch { /* storage unavailable */ }
  return 'a4';
}

export function setDefaultPageSize(k: PageSizeKey) {
  try { localStorage.setItem(DEFAULT_PAGE_KEY, k); } catch { /* storage unavailable */ }
}

function pageChoices(name: string, current: PageSizeKey | null): string {
  return (Object.keys(PAGE_SIZES) as PageSizeKey[]).map((k) => {
    const p = PAGE_SIZES[k];
    const w = 44, h = Math.round((w * p.cy) / p.cx);
    const box = h > w ? `width:${Math.round((w * w) / h)}px;height:${w}px` : `width:${w}px;height:${h}px`;
    return `<label class="page-choice"><input type="radio" name="${name}" value="${k}"${k === current ? ' checked' : ''}>`
      + `<span class="page-icon"><span style="${box}"></span></span>`
      + `<span><b>${p.label}</b><br><small>${p.detail}</small></span></label>`;
  }).join('');
}

/** 페이지 설정: change the current document's page format. */
export function pageSetupDialog(current: PageSizeKey | null, apply: (k: PageSizeKey, mode: 'fit' | 'stretch' | 'none', makeDefault: boolean) => void) {
  const body = html(`<div class="page-setup">
    <div class="page-choices">${pageChoices('page', current)}</div>
    <div class="page-sub">내용 맞추기</div>
    <label><input type="radio" name="mode" value="fit" checked> 비율 유지하며 맞추기 (글자 크기도 함께 조절)</label>
    <label><input type="radio" name="mode" value="stretch"> 페이지에 맞게 늘리기 (글자 크기 유지)</label>
    <label><input type="radio" name="mode" value="none"> 위치와 크기 그대로 두기</label>
    <label class="page-default"><input type="checkbox" id="page-default"> 새 문서의 기본 형식으로 사용</label>
  </div>`);
  const val = (n: string) => body.querySelector<HTMLInputElement>(`input[name=${n}]:checked`)?.value;
  dialog('페이지 설정', body, [
    {
      label: '적용', primary: true, action: () => {
        const k = val('page') as PageSizeKey | undefined;
        if (!k) return false;
        apply(k, val('mode') as 'fit' | 'stretch' | 'none', body.querySelector<HTMLInputElement>('#page-default')!.checked);
      },
    },
    { label: '취소' },
  ]);
}

/** 옵션: preferences that apply to new documents. */
export function optionsDialog() {
  const body = html(`<div class="page-setup">
    <div class="page-sub">새 문서의 페이지 형식</div>
    <div class="page-choices">${pageChoices('page', defaultPageSize())}</div>
    <small style="color:var(--color-text-secondary)">지금 열린 문서는 바뀌지 않습니다. 열린 문서는 파일 → 페이지 설정에서 바꿉니다.</small>
  </div>`);
  dialog('옵션', body, [
    {
      label: '확인', primary: true, action: () => {
        const k = body.querySelector<HTMLInputElement>('input[name=page]:checked')?.value as PageSizeKey | undefined;
        if (k) { setDefaultPageSize(k); toast(`새 문서 형식: ${PAGE_SIZES[k].label}`); }
      },
    },
    { label: '취소' },
  ]);
}

/** Layout chooser with schematic previews of placeholder positions. */
export function layoutDialog(doc: PptxDocument, title: string, pick: (layoutPath: string) => void) {
  const grid = document.createElement('div');
  grid.className = 'layout-grid';
  const { cx, cy } = doc.slideSize;
  const w = 158, h = (w * cy) / cx;
  let close = () => {};
  for (const l of doc.layouts()) {
    const card = document.createElement('div');
    card.className = 'layout-card';
    const th = document.createElement('div');
    th.className = 'lthumb';
    th.style.height = h + 'px';
    const lay = doc.pkg.getXml(l.path);
    const masterPath = doc.pkg.rels(l.path).find((r) => r.type.endsWith('/slideMaster'))?.target;
    const master = masterPath ? doc.pkg.getXml(masterPath) : null;
    let svg = `<svg width="${w}" height="${h}" viewBox="0 0 ${cx} ${cy}" style="position:absolute;left:0;top:0">`;
    for (const s of shapeChildren(spTree(lay))) {
      const ph = placeholder(s);
      if (!ph || ['dt', 'ftr', 'sldNum'].includes(ph.getAttribute('type') || '')) continue;
      let x = readXfrm(xfrmEl(s));
      if (!x) { const m = findInherited(ph, master, true); x = m ? readXfrm(xfrmEl(m)) : null; }
      if (!x) continue;
      const isTitle = ['title', 'ctrTitle'].includes(ph.getAttribute('type') || '');
      svg += `<rect x="${x.x}" y="${x.y}" width="${x.w}" height="${x.h}" fill="${isTitle ? '#dbe7fb' : '#f1f3f6'}" stroke="#8b929c" stroke-width="${cx / 300}" stroke-dasharray="${cx / 100} ${cx / 150}"/>`;
    }
    svg += '</svg>';
    th.innerHTML = svg;
    const name = document.createElement('div');
    name.textContent = l.name;
    card.append(th, name);
    card.onclick = () => { close(); pick(l.path); };
    grid.appendChild(card);
  }
  close = dialog(title, grid, [{ label: '취소' }]);
}

export interface PropsResult { x: Xfrm; fill?: string | null; line?: string | null; lineW?: number }

export function objectPropsDialog(x: Xfrm, apply: (r: { x: Xfrm; fill: string | null | undefined; line: string | null | undefined; lineW: number | undefined }) => void) {
  const cm = (v: number) => (v / 360000).toFixed(2);
  const body = html(`
    <div class="form">
      <label>가로 위치 (cm)</label><input type="number" step="0.01" name="x" value="${cm(x.x)}">
      <label>세로 위치 (cm)</label><input type="number" step="0.01" name="y" value="${cm(x.y)}">
      <label>너비 (cm)</label><input type="number" step="0.01" min="0" name="w" value="${cm(x.w)}">
      <label>높이 (cm)</label><input type="number" step="0.01" min="0" name="h" value="${cm(x.h)}">
      <label>회전 (°)</label><input type="number" step="1" name="rot" value="${x.rot}">
      <label></label><span></span>
      <label>채우기</label><span><input type="color" name="fill" value="#4472c4"> <label><input type="checkbox" name="fillOn"> 변경</label> <label><input type="checkbox" name="fillNone"> 없음</label></span>
      <label>선</label><span><input type="color" name="line" value="#2f528f"> <label><input type="checkbox" name="lineOn"> 변경</label> <label><input type="checkbox" name="lineNone"> 없음</label></span>
      <label>선 굵기 (pt)</label><input type="number" step="0.25" min="0" name="lineW" placeholder="변경 안 함">
    </div>`);
  const q = (n: string) => body.querySelector(`[name="${n}"]`) as HTMLInputElement;
  dialog('개체 속성', body, [
    { label: '취소' },
    {
      label: '적용', primary: true, action: () => {
        const nx: Xfrm = { ...x, x: Number(q('x').value) * 360000, y: Number(q('y').value) * 360000, w: Number(q('w').value) * 360000, h: Number(q('h').value) * 360000, rot: Number(q('rot').value) || 0 };
        const fill = q('fillNone').checked ? null : q('fillOn').checked ? q('fill').value.slice(1).toUpperCase() : undefined;
        const line = q('lineNone').checked ? null : q('lineOn').checked ? q('line').value.slice(1).toUpperCase() : undefined;
        const lw = q('lineW').value;
        apply({ x: nx, fill, line, lineW: lw ? Number(lw) : undefined });
      },
    },
  ]);
}

export function toast(msg: string) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2200);
}

export interface MenuItem { label?: string; shortcut?: string; disabled?: boolean; action?: () => void; sep?: boolean }

let openMenu: HTMLElement | null = null;

export function closeContextMenu() {
  openMenu?.remove();
  openMenu = null;
}

/** HOP-style dropdown used as a right-click menu. */
export function contextMenu(x: number, y: number, items: MenuItem[]) {
  closeContextMenu();
  const m = document.createElement('div');
  m.className = 'menu-dropdown ctx-menu';
  for (const it of items) {
    if (it.sep) { const s = document.createElement('div'); s.className = 'md-sep'; m.appendChild(s); continue; }
    const row = document.createElement('div');
    row.className = 'md-item' + (it.disabled ? ' disabled' : '');
    row.innerHTML = `<span class="md-label"></span><span class="md-shortcut"></span>`;
    row.querySelector('.md-label')!.textContent = it.label ?? '';
    row.querySelector('.md-shortcut')!.textContent = it.shortcut ?? '';
    row.addEventListener('mousedown', (e) => e.preventDefault());
    row.addEventListener('click', () => { closeContextMenu(); it.action?.(); });
    m.appendChild(row);
  }
  document.body.appendChild(m);
  m.style.left = Math.min(x, window.innerWidth - m.offsetWidth - 4) + 'px';
  m.style.top = Math.min(y, window.innerHeight - m.offsetHeight - 4) + 'px';
  openMenu = m;
  setTimeout(() => {
    const off = (e: Event) => {
      if (e instanceof MouseEvent && m.contains(e.target as Node)) return;
      closeContextMenu();
      document.removeEventListener('mousedown', off, true);
      window.removeEventListener('blur', off);
    };
    document.addEventListener('mousedown', off, true);
    window.addEventListener('blur', off);
  });
}
