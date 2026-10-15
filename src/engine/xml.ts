// XML helpers for OOXML (DrawingML / PresentationML) documents.

export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

type Prefix = keyof typeof NS;

function split(qname: string): [string, string] {
  const i = qname.indexOf(':');
  return [NS[qname.slice(0, i) as Prefix], qname.slice(i + 1)];
}

export function parseXml(text: string): XMLDocument {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('XML parse error: ' + err.textContent);
  return doc;
}

export function serializeXml(doc: XMLDocument): string {
  let s = new XMLSerializer().serializeToString(doc);
  // Imported fragments repeat the root's namespace declarations; drop the redundant copies.
  const rootEnd = s.indexOf('>', s.indexOf('<' + doc.documentElement.tagName)) + 1;
  if (rootEnd > 0) {
    let body = s.slice(rootEnd);
    for (const a of Array.from(doc.documentElement.attributes)) {
      if (a.name.startsWith('xmlns')) body = body.split(` ${a.name}="${a.value}"`).join('');
    }
    s = s.slice(0, rootEnd) + body;
  }
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}

/** First direct child matching a prefixed name such as "a:xfrm". */
export function child(el: Element | null | undefined, qname: string): Element | null {
  if (!el) return null;
  const [ns, local] = split(qname);
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === local && c.namespaceURI === ns) return c;
  }
  return null;
}

export function children(el: Element | null | undefined, qname?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  const [ns, local] = qname ? split(qname) : ['', ''];
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (!qname || (c.localName === local && c.namespaceURI === ns)) out.push(c);
  }
  return out;
}

/** Walk a path of direct children: path(el, 'p:spPr', 'a:xfrm', 'a:off'). */
export function path(el: Element | null | undefined, ...qnames: string[]): Element | null {
  let cur: Element | null = el ?? null;
  for (const q of qnames) cur = child(cur, q);
  return cur;
}

export function descendants(el: Element | Document, qname: string): Element[] {
  const [ns, local] = split(qname);
  return Array.from(el.getElementsByTagNameNS(ns, local));
}

export function create(doc: Document, qname: string, attrs: Record<string, string | number> = {}): Element {
  const [ns] = split(qname);
  const el = doc.createElementNS(ns, qname);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}

/** Return child qname, creating (appended or at the given position) if absent. */
export function ensure(el: Element, qname: string, before?: string[]): Element {
  const found = child(el, qname);
  if (found) return found;
  const made = create(el.ownerDocument, qname);
  let ref: Element | null = null;
  if (before) {
    for (const b of before) {
      ref = child(el, b);
      if (ref) break;
    }
  }
  el.insertBefore(made, ref);
  return made;
}

export function num(el: Element | null | undefined, attr: string, dflt = 0): number {
  const v = el?.getAttribute(attr);
  if (v == null || v === '') return dflt;
  const n = Number(v);
  return Number.isFinite(n) ? n : dflt;
}

export function attr(el: Element | null | undefined, name: string): string | null {
  return el?.getAttribute(name) ?? null;
}

export function getRId(el: Element | null | undefined, name: string): string | null {
  return el?.getAttributeNS(NS.r, name) || null;
}

export function removeChildren(el: Element, qname: string) {
  for (const c of children(el, qname)) el.removeChild(c);
}
