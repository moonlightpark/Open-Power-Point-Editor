// @vitest-environment jsdom
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

const dir = join(process.cwd(), 'tests');
import { describe, expect, it } from 'vitest';
import { PptxDocument, shapeChildren, spTree, placeholder } from '../src/engine/document';
import { ImageCache, renderSlide } from '../src/engine/render';
import { descendants } from '../src/engine/xml';

const fixture = () => new Uint8Array(readFileSync(join(dir, 'fixtures/sample.pptx')));
const texts = (doc: PptxDocument, i: number) => descendants(doc.pkg.getXml(doc.slides()[i].path)!, 'a:t').map((t) => t.textContent);

async function reload(doc: PptxDocument) {
  return PptxDocument.load(await doc.saveBytes());
}

describe('PptxDocument', () => {
  it('loads slides, layouts and size', async () => {
    const doc = await PptxDocument.load(fixture());
    expect(doc.slides()).toHaveLength(3);
    expect(doc.layouts().length).toBe(11);
    expect(doc.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
    expect(texts(doc, 0)).toContain('Opptx 샘플 발표');
    expect(doc.notesText(doc.slides()[2].path)).toBe('발표자 노트 예시');
  });

  it('resolves placeholder positions through the layout', async () => {
    const doc = await PptxDocument.load(fixture());
    const s = doc.slides()[0];
    const ctx = doc.context(s.path);
    const title = shapeChildren(spTree(ctx.slide)).find((x) => placeholder(x)?.getAttribute('type') === 'ctrTitle')!;
    const x = doc.effectiveXfrm(ctx, title)!;
    expect(x.w).toBeGreaterThan(0);
    expect(x.x).toBeGreaterThan(0);
  });

  it('adds, duplicates, moves and deletes slides with undo', async () => {
    const doc = await PptxDocument.load(fixture());
    doc.addSlide(doc.layouts()[1].path, 3);
    expect(doc.slides()).toHaveLength(4);
    doc.duplicateSlide(0);
    expect(doc.slides()).toHaveLength(5);
    expect(texts(doc, 1)).toEqual(texts(doc, 0));
    doc.moveSlide(0, 4);
    expect(texts(doc, 4)).toContain('Opptx 샘플 발표');
    doc.deleteSlides([4]);
    expect(doc.slides()).toHaveLength(4);
    doc.undo(); doc.undo(); doc.undo(); doc.undo();
    expect(doc.slides()).toHaveLength(3);
    expect(texts(doc, 0)).toContain('Opptx 샘플 발표');
    doc.redo();
    expect(doc.slides()).toHaveLength(4);
    const re = await reload(doc);
    expect(re.slides()).toHaveLength(4);
  });

  it('inserts shapes, text boxes, tables and pictures that survive save', async () => {
    const doc = await PptxDocument.load(fixture());
    const p = doc.slides()[1].path;
    const box = { x: 100000, y: 100000, w: 2000000, h: 1000000, rot: 0, flipH: false, flipV: false };
    doc.addShape(p, 'star5', box, '별');
    doc.addTextBox(p, box, '안녕하세요 & <test>');
    doc.addTable(p, 2, 3, box);
    doc.addPicture(p, new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), 'image/png', box);
    const re = await reload(doc);
    const xml = new XMLSerializer().serializeToString(re.pkg.getXml(re.slides()[1].path)!);
    expect(xml).toContain('prst="star5"');
    expect(xml).toContain('안녕하세요 &amp; &lt;test&gt;');
    expect(xml).toContain('<a:tbl>');
    expect(re.pkg.rels(re.slides()[1].path).some((r) => r.type.endsWith('/image'))).toBe(true);
  });

  it('formats runs and paragraphs, moves and deletes shapes', async () => {
    const doc = await PptxDocument.load(fixture());
    const p = doc.slides()[2].path;
    const shapes = shapeChildren(spTree(doc.pkg.getXml(p)));
    const rect = shapes.find((s) => descendants(s, 'a:t').some((t) => t.textContent === '둥근 사각형'))!;
    const id = rect.getElementsByTagNameNS('*', 'cNvPr')[0].getAttribute('id')!;
    doc.applyRunProps(p, [id], { b: true, sz: 32, color: '00FF00', font: '맑은 고딕' });
    doc.applyParaProps(p, [id], { algn: 'r' });
    doc.setXfrms(p, [{ id, x: { x: 1, y: 2, w: 3, h: 4, rot: 45, flipH: true, flipV: false } }]);
    doc.setFill(p, [id], 'FF0000');
    doc.setLine(p, [id], null);
    const rPr = descendants(rect, 'a:rPr')[0];
    expect(rPr.getAttribute('b')).toBe('1');
    expect(rPr.getAttribute('sz')).toBe('3200');
    expect(descendants(rect, 'a:pPr')[0].getAttribute('algn')).toBe('r');
    const re = await reload(doc);
    const xml = new XMLSerializer().serializeToString(re.pkg.getXml(p)!);
    expect(xml).toMatch(/<a:xfrm rot="2700000" flipH="1"><a:off x="1" y="2"\/><a:ext cx="3" cy="4"\/>/);
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
    doc.deleteShapes(p, [id]);
    expect(doc.findShape(p, id)).toBeNull();
    doc.undo();
    expect(doc.findShape(p, id)).not.toBeNull();
  });

  it('copies shapes with image relationships between slides', async () => {
    const doc = await PptxDocument.load(fixture());
    const from = doc.slides()[2].path, to = doc.slides()[0].path;
    const pic = descendants(doc.pkg.getXml(from)!, 'p:pic')[0];
    const id = pic.getElementsByTagNameNS('*', 'cNvPr')[0].getAttribute('id')!;
    const clip = doc.copyShapes(from, [id]);
    const [newId] = doc.pasteShapes(to, clip, 0);
    const pasted = doc.findShape(to, newId)!;
    const rId = descendants(pasted, 'a:blip')[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
    expect(doc.pkg.rel(to, rId)?.target).toMatch(/^ppt\/media\//);
  });

  it('copies, cuts and pastes slides (also after the source is deleted)', async () => {
    const doc = await PptxDocument.load(fixture());
    const clip = doc.copySlides([2, 0]);
    expect(clip).toHaveLength(2);
    doc.deleteSlides([0, 2]);
    expect(doc.slides()).toHaveLength(1);
    const at = doc.pasteSlides(clip, 1);
    expect(at).toBe(1);
    expect(doc.slides()).toHaveLength(3);
    expect(texts(doc, 1)).toContain('Opptx 샘플 발표');
    expect(texts(doc, 2)).toContain('도형과 표');
    // The pasted picture still resolves to its media part.
    const pic = descendants(doc.pkg.getXml(doc.slides()[2].path)!, 'a:blip')[0];
    const rId = pic.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
    expect(doc.pkg.rel(doc.slides()[2].path, rId)?.target).toMatch(/^ppt\/media\//);
    // Notes are not shared between slides.
    expect(doc.pkg.rels(doc.slides()[2].path).some((r) => r.type.endsWith('/notesSlide'))).toBe(false);
    const dup = doc.duplicateSlides([1, 2]);
    expect(dup).toBe(3);
    expect(doc.slides()).toHaveLength(5);
    expect(texts(doc, 3)).toEqual(texts(doc, 1));
    doc.undo();
    expect(doc.slides()).toHaveLength(3);
    const re = await reload(doc);
    expect(re.slides()).toHaveLength(3);
  });

  it('find and replace across slides', async () => {
    const doc = await PptxDocument.load(fixture());
    expect(doc.findAll('도형', false).length).toBeGreaterThan(0);
    expect(doc.replaceAll('PPTX', 'PowerPoint', true)).toBe(2);
    expect(texts(doc, 1)).toContain('PowerPoint로 저장');
  });

  it('renders every slide to DOM', async () => {
    const doc = await PptxDocument.load(fixture());
    const images = new ImageCache(doc);
    for (const s of doc.slides()) {
      const r = renderSlide(doc, s.path, images, { editor: true });
      expect(r.width).toBe(960);
      expect(r.shapes.length).toBeGreaterThan(0);
    }
    const r = renderSlide(doc, doc.slides()[2].path, images);
    expect(r.el.querySelector('table.tbl')).not.toBeNull();
    expect(r.el.querySelectorAll('td[data-cell]')).toHaveLength(9);
    expect(r.el.textContent).toContain('굵은 빨간 글자');
    const red = Array.from(r.el.querySelectorAll('span')).find((x) => x.textContent === '굵은 빨간 글자') as HTMLElement;
    expect(red.style.fontWeight).toBe('bold');
    expect(red.style.color).toBe('rgb(192, 0, 0)');
    expect(red.style.fontSize).toBe('24px');
    const body = renderSlide(doc, doc.slides()[1].path, images);
    expect(body.el.querySelectorAll('.tx-p.has-bullet').length).toBe(5);
  });

  it('writes an edited deck for external validation', async () => {
    const doc = await PptxDocument.load(fixture());
    const p = doc.slides()[0].path;
    doc.addShape(p, 'ellipse', { x: 500000, y: 500000, w: 1000000, h: 1000000, rot: 0, flipH: false, flipV: false }, '타원');
    doc.addSlide(doc.layouts()[1].path, 1);
    doc.duplicateSlide(2);
    doc.deleteSlides([3]);
    const bytes = await doc.saveBytes();
    mkdirSync(join(dir, 'out'), { recursive: true });
    writeFileSync(join(dir, 'out/edited.pptx'), bytes);
  });
});
