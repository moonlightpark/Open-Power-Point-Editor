// Inline SVG toolbar icons (stroke-based, 24x24).
const S = (body: string) => `<svg viewBox="0 0 24 24" fill="none" stroke="#3b4452" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;

export const ICONS: Record<string, string> = {
  open: S('<path d="M3 7h6l2 2h10v10H3z" fill="#f6d68a" stroke="#b68b21"/><path d="M3 9h18" stroke="#b68b21"/>'),
  save: S('<path d="M4 4h13l3 3v13H4z" fill="#dbe7fb" stroke="#2b6fd6"/><rect x="7" y="4" width="9" height="5" stroke="#2b6fd6"/><rect x="7" y="13" width="10" height="7" fill="#fff" stroke="#2b6fd6"/>'),
  cut: S('<circle cx="6" cy="17" r="3"/><circle cx="18" cy="17" r="3"/><path d="M8 15 18 3M16 15 6 3"/>'),
  copy: S('<rect x="8" y="8" width="12" height="13" rx="1" fill="#fff"/><path d="M16 8V4H4v13h4"/>'),
  paste: S('<rect x="5" y="4" width="14" height="17" rx="1" fill="#f3e3c3" stroke="#a07a3a"/><rect x="9" y="2.5" width="6" height="3" rx="1" fill="#fff"/><rect x="9" y="9" width="11" height="12" fill="#fff"/>'),
  undo: S('<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>'),
  redo: S('<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>'),
  slide: S('<rect x="3" y="5" width="18" height="13" rx="1" fill="#fff"/><path d="M6 9h8M6 12h12M6 15h9" stroke="#8a93a3"/><circle cx="19" cy="19" r="4" fill="#2b6fd6" stroke="none"/><path d="M19 17v4M17 19h4" stroke="#fff"/>'),
  layout: S('<rect x="3" y="4" width="18" height="16" rx="1" fill="#fff"/><path d="M3 9h18M10 9v11"/>'),
  textbox: S('<rect x="3" y="4" width="18" height="16" rx="1" stroke-dasharray="2 2"/><path d="M8 8h8M12 8v9"/>'),
  shape: S('<circle cx="9" cy="9" r="5" fill="#cfe0fb" stroke="#2b6fd6"/><rect x="10" y="11" width="10" height="9" fill="#f9d7b5" stroke="#d9772b"/>'),
  image: S('<rect x="3" y="5" width="18" height="14" rx="1" fill="#fff"/><circle cx="8.5" cy="9.5" r="1.8" fill="#f2b632" stroke="none"/><path d="m3 17 5-5 4 4 3-3 6 6" stroke="#2f9e5b"/>'),
  table: S('<rect x="3" y="4" width="18" height="16" rx="1" fill="#fff"/><path d="M3 9h18M3 14.5h18M9 4v16M15 4v16"/><path d="M3 4h18v5H3z" fill="#cfe0fb"/>'),
  fill: S('<path d="m5 12 7-7 7 7-7 7z" fill="#fff"/><path d="M19 15s2 2.3 2 3.5a2 2 0 0 1-4 0c0-1.2 2-3.5 2-3.5z" fill="#2b6fd6" stroke="#2b6fd6"/><rect x="3" y="21" width="18" height="2" fill="#f2b632" stroke="none"/>'),
  line: S('<path d="M4 18 20 6" stroke-width="2.2"/><rect x="3" y="21" width="18" height="2" fill="#2b6fd6" stroke="none"/>'),
  props: S('<rect x="3" y="4" width="18" height="16" rx="1" fill="#fff"/><path d="M7 9h6M7 13h10M7 17h4"/><circle cx="17" cy="9" r="1.5"/>'),
  front: S('<rect x="3" y="3" width="11" height="11" fill="#e3e6eb"/><rect x="9" y="9" width="12" height="12" fill="#cfe0fb" stroke="#2b6fd6"/>'),
  back: S('<rect x="9" y="9" width="12" height="12" fill="#cfe0fb" stroke="#2b6fd6"/><rect x="3" y="3" width="11" height="11" fill="#e3e6eb"/>'),
  find: S('<circle cx="10" cy="10" r="6" fill="#fff"/><path d="m15 15 6 6" stroke-width="2.2"/>'),
  play: S('<rect x="3" y="4" width="18" height="13" rx="1" fill="#fff"/><path d="m10 7.5 5 3-5 3z" fill="#2b6fd6" stroke="#2b6fd6"/><path d="M12 17v3M8 21h8"/>'),
  al: S('<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>'),
  ac: S('<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>'),
  ar: S('<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>'),
  aj: S('<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>'),
};

export function installIcons(root: ParentNode = document) {
  root.querySelectorAll<HTMLElement>('[data-icon]').forEach((el) => {
    const svg = ICONS[el.dataset.icon!];
    if (svg) el.innerHTML = svg;
  });
}
