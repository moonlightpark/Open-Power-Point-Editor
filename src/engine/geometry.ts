// Preset (prstGeom) and custom (custGeom) shape outlines as SVG path data.
import { child, children, num } from './xml';

export interface PresetInfo { id: string; label: string }

/** Shapes offered in the insert menu. */
export const INSERTABLE: PresetInfo[] = [
  { id: 'rect', label: '사각형' },
  { id: 'roundRect', label: '둥근 사각형' },
  { id: 'ellipse', label: '타원' },
  { id: 'triangle', label: '삼각형' },
  { id: 'rtTriangle', label: '직각 삼각형' },
  { id: 'diamond', label: '마름모' },
  { id: 'parallelogram', label: '평행사변형' },
  { id: 'trapezoid', label: '사다리꼴' },
  { id: 'pentagon', label: '오각형' },
  { id: 'hexagon', label: '육각형' },
  { id: 'octagon', label: '팔각형' },
  { id: 'star5', label: '별' },
  { id: 'rightArrow', label: '오른쪽 화살표' },
  { id: 'leftArrow', label: '왼쪽 화살표' },
  { id: 'upArrow', label: '위쪽 화살표' },
  { id: 'downArrow', label: '아래쪽 화살표' },
  { id: 'chevron', label: '갈매기형' },
  { id: 'homePlate', label: '홈 플레이트' },
  { id: 'plus', label: '더하기' },
  { id: 'wedgeRectCallout', label: '말풍선' },
  { id: 'line', label: '선' },
];

function adj(gd: Map<string, number>, name: string, dflt: number) {
  return gd.has(name) ? gd.get(name)! : dflt;
}

function poly(pts: [number, number][]) {
  return 'M' + pts.map((p) => p[0].toFixed(2) + ',' + p[1].toFixed(2)).join(' L') + ' Z';
}

function regular(n: number, w: number, h: number, start = -Math.PI / 2): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < n; i++) {
    const a = start + (i * 2 * Math.PI) / n;
    pts.push([w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)]);
  }
  return poly(pts);
}

export function presetPath(prst: string, w: number, h: number, avLst: Element | null): string {
  const gd = new Map<string, number>();
  for (const g of children(avLst, 'a:gd')) {
    const m = (g.getAttribute('fmla') || '').match(/val\s+(-?\d+)/);
    if (m) gd.set(g.getAttribute('name') || '', Number(m[1]));
  }
  const ss = Math.min(w, h);
  switch (prst) {
    case 'ellipse': case 'flowChartConnector':
      return `M0,${h / 2} A${w / 2},${h / 2} 0 1 0 ${w},${h / 2} A${w / 2},${h / 2} 0 1 0 0,${h / 2} Z`;
    case 'roundRect': case 'flowChartAlternateProcess': {
      const r = Math.min((ss * adj(gd, 'adj', 16667)) / 100000, w / 2, h / 2);
      return `M${r},0 H${w - r} A${r},${r} 0 0 1 ${w},${r} V${h - r} A${r},${r} 0 0 1 ${w - r},${h} H${r} A${r},${r} 0 0 1 0,${h - r} V${r} A${r},${r} 0 0 1 ${r},0 Z`;
    }
    case 'triangle': case 'flowChartExtract': {
      const x = (w * adj(gd, 'adj', 50000)) / 100000;
      return poly([[x, 0], [w, h], [0, h]]);
    }
    case 'rtTriangle': return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond': case 'flowChartDecision': return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': case 'flowChartInputOutput': {
      const x = (ss * adj(gd, 'adj', 25000)) / 100000;
      return poly([[x, 0], [w, 0], [w - x, h], [0, h]]);
    }
    case 'trapezoid': {
      const x = (ss * adj(gd, 'adj', 25000)) / 100000;
      return poly([[x, 0], [w - x, 0], [w, h], [0, h]]);
    }
    case 'pentagon': return regular(5, w, h);
    case 'hexagon': {
      const x = (ss * adj(gd, 'adj', 25000)) / 100000;
      return poly([[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]]);
    }
    case 'octagon': {
      const x = (ss * adj(gd, 'adj', 29289)) / 100000;
      return poly([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]]);
    }
    case 'star5': {
      const pts: [number, number][] = [];
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const k = i % 2 ? 0.382 : 1;
        pts.push([w / 2 + (w / 2) * k * Math.cos(a), h / 2 + (h / 2) * k * Math.sin(a)]);
      }
      return poly(pts);
    }
    case 'rightArrow': case 'leftArrow': {
      const a1 = (h * adj(gd, 'adj1', 50000)) / 100000;
      const hd = Math.min((ss * adj(gd, 'adj2', 50000)) / 100000, w);
      const y1 = (h - a1) / 2, y2 = y1 + a1;
      const pts: [number, number][] = [[0, y1], [w - hd, y1], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, y2], [0, y2]];
      return poly(prst === 'leftArrow' ? pts.map(([x, y]) => [w - x, y]) : pts);
    }
    case 'upArrow': case 'downArrow': {
      const a1 = (w * adj(gd, 'adj1', 50000)) / 100000;
      const hd = Math.min((ss * adj(gd, 'adj2', 50000)) / 100000, h);
      const x1 = (w - a1) / 2, x2 = x1 + a1;
      const pts: [number, number][] = [[x1, h], [x1, hd], [0, hd], [w / 2, 0], [w, hd], [x2, hd], [x2, h]];
      return poly(prst === 'downArrow' ? pts.map(([x, y]) => [x, h - y]) : pts);
    }
    case 'chevron': {
      const x = Math.min((ss * adj(gd, 'adj', 50000)) / 100000, w);
      return poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]);
    }
    case 'homePlate': {
      const x = Math.min((ss * adj(gd, 'adj', 50000)) / 100000, w);
      return poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h]]);
    }
    case 'plus': case 'flowChartSummingJunction': {
      const x = (ss * adj(gd, 'adj', 25000)) / 100000;
      return poly([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]]);
    }
    case 'wedgeRectCallout': case 'wedgeRoundRectCallout': {
      const tx = w / 2 + (w * adj(gd, 'adj1', -20833)) / 100000;
      const ty = h / 2 + (h * adj(gd, 'adj2', 62500)) / 100000;
      return poly([[0, 0], [w, 0], [w, h], [w * 0.583, h], [tx, ty], [w * 0.333, h], [0, h]]);
    }
    case 'line': case 'straightConnector1': case 'bentConnector3': case 'curvedConnector3':
      return `M0,0 L${w},${h}`;
    case 'flowChartTerminator': {
      const r = Math.min(h / 2, w / 2);
      return `M${r},0 H${w - r} A${r},${h / 2} 0 0 1 ${w - r},${h} H${r} A${r},${h / 2} 0 0 1 ${r},0 Z`;
    }
    default:
      return `M0,0 H${w} V${h} H0 Z`;
  }
}

export const OPEN_PRESETS = new Set(['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3', 'arc']);

/** Convert a:custGeom into a path scaled into a w x h box. */
export function customPath(custGeom: Element, w: number, h: number): string {
  let d = '';
  for (const p of children(child(custGeom, 'a:pathLst'), 'a:path')) {
    const pw = num(p, 'w', 0) || w, ph = num(p, 'h', 0) || h;
    const sx = w / pw, sy = h / ph;
    let cx = 0, cy = 0;
    const pt = (e: Element | null): [number, number] => [num(e, 'x') * sx, num(e, 'y') * sy];
    for (const c of children(p)) {
      const pts = children(c, 'a:pt').map(pt);
      switch (c.localName) {
        case 'moveTo': [cx, cy] = pts[0]; d += `M${cx},${cy} `; break;
        case 'lnTo': [cx, cy] = pts[0]; d += `L${cx},${cy} `; break;
        case 'cubicBezTo': d += `C${pts.map((q) => q.join(',')).join(' ')} `; [cx, cy] = pts[2]; break;
        case 'quadBezTo': d += `Q${pts.map((q) => q.join(',')).join(' ')} `; [cx, cy] = pts[1]; break;
        case 'arcTo': {
          const wR = num(c, 'wR') * sx, hR = num(c, 'hR') * sy;
          const st = (num(c, 'stAng') / 60000) * (Math.PI / 180);
          const sw = (num(c, 'swAng') / 60000) * (Math.PI / 180);
          const ox = cx - wR * Math.cos(st), oy = cy - hR * Math.sin(st);
          const ex = ox + wR * Math.cos(st + sw), ey = oy + hR * Math.sin(st + sw);
          d += `A${wR},${hR} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${ex},${ey} `;
          cx = ex; cy = ey;
          break;
        }
        case 'close': d += 'Z '; break;
      }
    }
  }
  return d || `M0,0 H${w} V${h} H0 Z`;
}
