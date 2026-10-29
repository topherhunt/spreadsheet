import { Run, RunFont, Workbook } from '../../core/model';
import { BorderSide, fontName, resolveColor } from '../../core/styles';

export const dark = window.matchMedia('(prefers-color-scheme: dark)');

function luminance(hex: string): number {
  const n = parseInt(hex.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => { const s = c / 255; return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * In dark mode, black text and white fills mean "the default" in practice
 * (reports set them explicitly), so they follow the theme instead.
 */
export function textColor(hex: string | undefined): string | undefined {
  if (!hex) return undefined;
  if (dark.matches && luminance(hex) < 0.02) return undefined;
  return hex;
}

function fillColor(hex: string | undefined): string | undefined {
  if (!hex) return undefined;
  if (dark.matches && luminance(hex) > 0.95) return undefined;
  return hex;
}

export function cssFont(name: string): string {
  return `"${name.replace(/"/g, '')}", "Carlito", "Liberation Sans", Arial, sans-serif`;
}

const BORDER_WIDTH: Record<string, string> = {
  thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed',
  medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed',
  thick: '3px solid', double: '3px double',
};

function borderCss(wb: Workbook, side: BorderSide | undefined): string | undefined {
  if (!side || side.style === 'none') return undefined;
  const color = resolveColor(wb.styles, side.color);
  return `${BORDER_WIDTH[side.style] ?? '1px solid'} ${textColor(color) ?? 'var(--fg)'}`;
}

export interface XfInfo {
  cls: string;
  hAlign?: string;
  vAlign: string;
  wrap: boolean;
  /** The cell's fill paints over gridlines. */
  filled: boolean;
  bordered: boolean;
  fontSizePt: number;
  bold: boolean; italic: boolean; underline: boolean; strike: boolean;
  fontName: string;
  color?: string;
  /** The text color actually drawn (after dark-mode adaptation); unset means the theme's. */
  fg?: string;
  fill?: string;
  indent: number;
  rotation: number;
}

/**
 * Generates one CSS class per cell style on demand, scaled to the zoom level.
 * The fill and borders go on the cell box; font rules on its text.
 */
export class StyleCache {
  private info = new Map<number, XfInfo>();
  private sheet: CSSStyleSheet;

  constructor(private wb: Workbook, private zoom: number) {
    this.sheet = new CSSStyleSheet();
    document.adoptedStyleSheets = [...document.adoptedStyleSheets.filter((s) => !(s as any).__cellStyles), this.sheet];
    (this.sheet as any).__cellStyles = true;
  }

  get(xf: number | undefined): XfInfo {
    const id = xf ?? 0;
    let info = this.info.get(id);
    if (!info) { info = this.build(id); this.info.set(id, info); }
    return info;
  }

  private build(id: number): XfInfo {
    const ss = this.wb.styles;
    const xf = ss.xfs[id];
    if (!xf) throw new Error(`Cell style ${id} is out of range`);
    const font = ss.fonts[xf.fontId];
    const fill = ss.fills[xf.fillId];
    const border = ss.borders[xf.borderId];
    const name = fontName(ss, font);
    const color = textColor(resolveColor(ss, font.color));
    let bg: string | undefined;
    if (fill && fill.pattern !== 'none' && !fill.gradient) bg = fillColor(resolveColor(ss, fill.pattern === 'solid' ? fill.fg : fill.fg ?? fill.bg));
    const cls = `x${id}`;
    const box: string[] = [];
    if (bg) box.push(`background:${bg}`);
    const sides = { left: borderCss(this.wb, border?.left), right: borderCss(this.wb, border?.right), top: borderCss(this.wb, border?.top), bottom: borderCss(this.wb, border?.bottom) };
    for (const [side, css] of Object.entries(sides)) if (css) box.push(`border-${side}:${css}`);
    const text: string[] = [
      `font-family:${cssFont(name)}`,
      `font-size:${(font.size * this.zoom).toFixed(2)}pt`,
    ];
    if (font.bold) text.push('font-weight:700');
    if (font.italic) text.push('font-style:italic');
    const deco = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) text.push(`text-decoration:${deco}`);
    // Auto-colored text on a light fill stays dark even in dark mode.
    const fg = color ?? (bg && luminance(bg) > 0.4 && dark.matches ? '#1f1f1f' : undefined);
    if (fg) text.push(`color:${fg}`);
    if (xf.align.indent) text.push(`padding-left:${3 + xf.align.indent * 9 * this.zoom}px`);
    this.sheet.insertRule(`.${cls}{${box.join(';')}}`, this.sheet.cssRules.length);
    this.sheet.insertRule(`.${cls}>.t{${text.join(';')}}`, this.sheet.cssRules.length);
    return {
      cls,
      hAlign: xf.align.horizontal,
      vAlign: xf.align.vertical ?? 'bottom',
      wrap: !!xf.align.wrapText,
      filled: !!bg,
      bordered: Object.values(sides).some(Boolean),
      fontSizePt: font.size,
      bold: font.bold, italic: font.italic, underline: !!font.underline, strike: font.strike,
      fontName: name,
      color: resolveColor(ss, font.color),
      fg,
      fill: bg,
      indent: xf.align.indent ?? 0,
      rotation: xf.align.textRotation ?? 0,
    };
  }
}

export function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
}

export function runStyle(font: RunFont | undefined, zoom: number): string {
  if (!font) return '';
  const out: string[] = [];
  if (font.bold !== undefined) out.push(`font-weight:${font.bold ? 700 : 400}`);
  if (font.italic !== undefined) out.push(`font-style:${font.italic ? 'italic' : 'normal'}`);
  if (font.underline !== undefined || font.strike !== undefined) {
    const d = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    out.push(`text-decoration:${d || 'none'}`);
  }
  const c = textColor(font.color);
  if (c) out.push(`color:${c}`);
  if (font.size) out.push(`font-size:${(font.size * zoom).toFixed(2)}pt`);
  if (font.name) out.push(`font-family:${cssFont(font.name)}`);
  return out.join(';');
}

export function runsHtml(runs: Run[], zoom: number): string {
  return runs.map((r) => {
    const style = runStyle(r.font, zoom);
    const cls = r.link ? ' class="lnk"' : '';
    return `<span${cls}${style ? ` style="${escapeHtml(style)}"` : ''}>${escapeHtml(r.text)}</span>`;
  }).join('');
}
