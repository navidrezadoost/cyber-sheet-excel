// Minimal import; we rely on runtime availability, typing via any to avoid dependency on @types
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import { unzipSync, strFromU8 } from 'fflate';
import { DEFAULT_WORKSHEET_COLS, DEFAULT_WORKSHEET_ROWS, Workbook, CellStyle } from '@cyber-sheet/core';

// Export new lightweight parser
export * from './import';
export * from './LightweightParser';
export * from './CommentParser';
export * from './export';
export * from './progressive';

type FetchLike = (input: RequestInfo, init?: RequestInit) => Promise<Response>;

function displayDimension(value: number, defaultValue: number): number {
  return Math.max(defaultValue, value);
}

export async function loadXlsxFromUrl(url: string, fetchFn?: FetchLike): Promise<Workbook> {
  const f = fetchFn ?? fetch;
  const res = await f(url);
  if (!res.ok) throw new Error(`Failed to fetch XLSX: ${res.status}`);
  const buf = new Uint8Array(await res.arrayBuffer());
  return loadXlsxFromArrayBuffer(buf);
}

export function loadXlsxFromArrayBuffer(data: Uint8Array): Workbook {
  const files = unzipSync(data);
  const getText = (name: string) => (files[name] ? strFromU8(files[name]) : undefined);
  const wb = new Workbook();

  // Theme colors (optional)
  const themeXml = getText('xl/theme/theme1.xml') || '';
  const themePalette = parseTheme(themeXml);

  // Parse workbook rels to find sheets
  const wbXml = getText('xl/workbook.xml');
  if (!wbXml) throw new Error('Invalid XLSX: missing workbook.xml');
  
  const sheetEntries = Array.from(wbXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="(\d+)"[^>]*r:id="([^"]+)"/g)).map((m) => {
    const mm = m as unknown as RegExpMatchArray; return { name: mm[1], id: mm[2], rid: mm[3] };
  });
  const relsXml = getText('xl/_rels/workbook.xml.rels') || '';
  
  // Fix: Parse Id and Target separately since they can appear in any order
  const rels = new Map(Array.from(relsXml.matchAll(/<Relationship[^>]*?\/>/g)).map((m) => {
    const rel = m[0];
    const idMatch = rel.match(/Id="([^"]+)"/);
    const targetMatch = rel.match(/Target="([^"]+)"/);
    if (idMatch && targetMatch) {
      return [idMatch[1], targetMatch[1]] as const;
    }
    return null;
  }).filter(Boolean) as Array<readonly [string, string]>);

  // Shared strings
  const sstXml = getText('xl/sharedStrings.xml') || '';
  const sharedStrings = Array.from(sstXml.matchAll(/<si>\s*<t[^>]*>([\s\S]*?)<\/t>\s*<\/si>/g)).map((m) => {
    const mm = m as unknown as RegExpMatchArray; return decodeXml(mm[1]);
  });

  // Styles (limited): number formats, fonts (bold/italic/size/color/name), fills
  const stylesXml = getText('xl/styles.xml') || '';
  const styles = parseStyles(stylesXml, themePalette);

  for (const s of sheetEntries) {
    const target = rels.get(s.rid);
    if (!target) continue;
    // Fix: handle both absolute (/xl/worksheets/sheet1.xml) and relative (worksheets/sheet1.xml) paths
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
    const sheetXml = getText(path);
    if (!sheetXml) continue;
    
    // Parse dimension to get declared sheet size, then scan cells as a fallback.
    const dimMatch = sheetXml.match(/<dimension[^>]*ref="([A-Z]+\d+:)?([A-Z]+\d+)"/);
    let maxRow = 1;
    let maxCol = 1;
    
    if (dimMatch) {
      const endRef = dimMatch[2]; // e.g., "Z100"
      const endAddr = a1ToAddress(endRef);
      maxRow = endAddr.row;
      maxCol = endAddr.col;
    }

    // Parse rows and c elements before creating the sheet so files with missing
    // or stale dimensions still get their full real size.
    // Example: <c r="A1" t="s"><v>0</v></c>
    const rowMatches = Array.from(sheetXml.matchAll(/<row[^>]*r="(\d+)"[\s\S]*?<\/row>/g)) as unknown as RegExpMatchArray[];
    for (const rowMatch of rowMatches) {
      const rowR = parseInt((rowMatch as any)[1], 10);
      maxRow = Math.max(maxRow, rowR);
      const rowXml = (rowMatch as any)[0] as string;
      const cellRefs = Array.from(rowXml.matchAll(/<c\s+[^>]*\br="([A-Z]+\d+)"/g)) as unknown as RegExpMatchArray[];
      for (const refMatch of cellRefs) {
        const addr = a1ToAddress((refMatch as any)[1] as string);
        maxRow = Math.max(maxRow, addr.row);
        maxCol = Math.max(maxCol, addr.col);
      }
    }

    const ws = wb.addSheet(
      s.name,
      displayDimension(maxRow, DEFAULT_WORKSHEET_ROWS),
      displayDimension(maxCol, DEFAULT_WORKSHEET_COLS),
    );
    
    for (const rowMatch of rowMatches) {
      const rowR = parseInt((rowMatch as any)[1], 10);
      const rowXml = (rowMatch as any)[0] as string;
      // Match both self-closing <c .../> and regular <c ...>...</c> tags
      // Simpler approach: match <c ...> and capture everything, then check for self-closing
      const cellMatches = Array.from(rowXml.matchAll(/<c\s+([^>]+?)(?:\/>|>(.*?)<\/c>)/g)) as unknown as RegExpMatchArray[];
      
      for (const cm of cellMatches) {
        const fullAttrs = (cm as any)[1] as string; // All attributes
        const inner = (cm as any)[2] as string || ''; // Inner content (empty for self-closing)
        
        // Extract r attribute manually
        const rMatch = fullAttrs.match(/r="([A-Z]+\d+)"/);
        const addr = rMatch ? rMatch[1] : undefined;
        
        if (!addr) continue; // Skip if we can't parse the address
        
        const tMatch = fullAttrs.match(/\bt="(\w+)"/);
        const t = tMatch ? tMatch[1] : '';
        const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
        const val = vMatch ? vMatch[1] : '';
        const sMatch = fullAttrs.match(/\bs="(\d+)"/);
        const sIdx = sMatch ? parseInt(sMatch[1], 10) : undefined;
        const mappedStyle = sIdx != null ? styles.xfStyles[sIdx] : undefined;
        const { col } = a1ToAddress(addr);
        
        // Only set value if cell has content (not self-closing)
        if (vMatch || inner.length > 0) {
          if (t === 's') {
            const idx = parseInt(val, 10);
            ws.setCellValue({ row: rowR, col }, sharedStrings[idx] ?? '');
          } else if (t === 'b') {
            ws.setCellValue({ row: rowR, col }, val === '1');
          } else if (t === 'str' || t === 'inlineStr') {
            ws.setCellValue({ row: rowR, col }, decodeXml(val));
          } else {
            // number or date (we'll treat as number for now; formatting later)
            const num = Number(val);
            ws.setCellValue({ row: rowR, col }, Number.isFinite(num) ? num : val);
          }
        }
        if (mappedStyle) {
          ws.setCellStyle({ row: rowR, col }, mappedStyle);
        }
      }
    }
  }
  
  return wb;
}

function a1ToAddress(a1: string): { row: number; col: number } {
  const m = a1.match(/^([A-Z]+)(\d+)$/);
  if (!m) return { row: 1, col: 1 };
  const letters = m[1];
  const row = parseInt(m[2], 10);
  let col = 0; for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return { row, col };
}

function decodeXml(s: string): string { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'); }

function parseStyles(xml: string, theme?: ThemePalette): { xfStyles: (CellStyle & { numberFormat?: string })[] } {
  if (!xml) return { xfStyles: [] };
  // numFmts
  const numFmtMap = new Map<number, string>();
  const numFmtMatches = Array.from(xml.matchAll(/<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]+)"/g)) as unknown as RegExpMatchArray[];
  for (const m of numFmtMatches) {
    const id = parseInt((m as any)[1], 10); const code = (m as any)[2] as string; numFmtMap.set(id, code);
  }
  // built-in minimal mapping
  const builtin = new Map<number, string>([
    [0, 'General'], [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'], [9, '0%'], [10, '0.00%'], [14, 'm/d/yy'], [22, 'm/d/yy h:mm']
  ]);

  // fonts
  const fonts: { 
    bold?: boolean; 
    italic?: boolean;
    underline?: boolean;
    size?: number; 
    color?: string; 
    name?: string;
    strikethrough?: boolean;
    superscript?: boolean;
    subscript?: boolean;
  }[] = [];
  const fontsXml = (xml.match(/<fonts[\s\S]*?<\/fonts>/) || [])[0] || '';
  const fontMatches = Array.from(fontsXml.matchAll(/<font>([\s\S]*?)<\/font>/g)) as unknown as RegExpMatchArray[];
  for (const fm of fontMatches) {
    const fxml = (fm as any)[1] as string;
    const bold = /<b\b/.test(fxml) ? true : undefined;
    const italic = /<i\b/.test(fxml) ? true : undefined;
    const underline = /<u\b/.test(fxml) ? true : undefined;
    const strikethrough = /<strike\b/.test(fxml) ? true : undefined; // Phase 1 UI
    
    // Phase 1 UI: superscript/subscript (mutually exclusive)
    const vertAlignMatch = fxml.match(/<vertAlign[^>]*val="(\w+)"/);
    const vertAlign = vertAlignMatch ? vertAlignMatch[1] : undefined;
    const superscript = vertAlign === 'superscript' ? true : undefined;
    const subscript = vertAlign === 'subscript' ? true : undefined;
    
    const szMatch = fxml.match(/<sz[^>]*val="([\d.]+)"/);
    const size = szMatch ? parseFloat(szMatch[1]) : undefined;
    const nameMatch = fxml.match(/<name[^>]*val="([^"]+)"/);
    const name = nameMatch ? nameMatch[1] : undefined;
    // color can be rgb or theme + optional tint
    let color: string | undefined;
    const colorRgb = fxml.match(/<color[^>]*rgb="([A-Fa-f0-9]{8})"/);
    const colorTheme = fxml.match(/<color[^>]*theme="(\d+)"/);
    const colorTint = fxml.match(/<color[^>]*tint="(-?[\d.]+)"/);
    if (colorRgb) color = argbToCss(colorRgb[1]);
    else if (colorTheme && theme) {
      const idx = parseInt(colorTheme[1], 10);
      const tint = colorTint ? parseFloat(colorTint[1]) : 0;
      color = themeColorToCss(theme, idx, tint);
    }
    fonts.push({ bold, italic, underline, strikethrough, superscript, subscript, size, color, name });
  }

  // fills
  const fills: { fill?: string }[] = [];
  const fillsXml = (xml.match(/<fills[\s\S]*?<\/fills>/) || [])[0] || '';
  const fillMatches = Array.from(fillsXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) as unknown as RegExpMatchArray[];
  for (const fl of fillMatches) {
    const fxml = (fl as any)[1] as string;
    let fill: string | undefined;
    // Check if pattern type is solid (other patterns like 'none' should not produce a fill)
    const patternType = fxml.match(/<patternFill[^>]*patternType="([^"]+)"/);
    const isSolid = patternType && patternType[1] !== 'none';
    
    if (isSolid) {
      const fgRgb = fxml.match(/<fgColor[^>]*rgb="([A-Fa-f0-9]{8})"/);
      const fgTheme = fxml.match(/<fgColor[^>]*theme="(\d+)"/);
      const fgTint = fxml.match(/<fgColor[^>]*tint="(-?[\d.]+)"/);
      if (fgRgb) fill = argbToCss(fgRgb[1]);
      else if (fgTheme && theme) {
        const idx = parseInt(fgTheme[1], 10);
        const tint = fgTint ? parseFloat(fgTint[1]) : 0;
        fill = themeColorToCss(theme, idx, tint);
      }
    }
    fills.push({ fill });
  }

  // borders
  const borders: { top?: string; right?: string; bottom?: string; left?: string }[] = [];
  const bordersXml = (xml.match(/<borders[\s\S]*?<\/borders>/) || [])[0] || '';
  const borderMatches = Array.from(bordersXml.matchAll(/<border>([\s\S]*?)<\/border>/g)) as unknown as RegExpMatchArray[];
  for (const bm of borderMatches) {
    const bxml = (bm as any)[1] as string;
    const getSide = (name: string) => {
      // rgb or theme
      const rgb = bxml.match(new RegExp(`<${name}[^>]*color[^>]*rgb="([A-Fa-f0-9]{8})"`));
      if (rgb) return argbToCss(rgb[1]);
      const th = bxml.match(new RegExp(`<${name}[^>]*color[^>]*theme="(\\d+)"`));
      if (th && theme) {
        const tIdx = parseInt(th[1], 10);
        const tintM = bxml.match(new RegExp(`<${name}[^>]*color[^>]*tint="(-?[\\d.]+)"`));
        const tint = tintM ? parseFloat(tintM[1]) : 0;
        return themeColorToCss(theme, tIdx, tint);
      }
      return undefined;
    };
    borders.push({ top: getSide('top'), right: getSide('right'), bottom: getSide('bottom'), left: getSide('left') });
  }

  // xfs
  const xfs: (CellStyle & { numberFormat?: string })[] = [];
  const xfsXml = (xml.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [])[0] || '';
  const xfMatches = Array.from(xfsXml.matchAll(/<xf[^>]*>/g)) as unknown as RegExpMatchArray[];
  for (const xm of xfMatches) {
    const tag = (xm as any)[0] as string;
    const numFmtId = parseInt((tag.match(/numFmtId="(\d+)"/) || [])[1] || '0', 10);
    const fontId = parseInt((tag.match(/fontId="(\d+)"/) || [])[1] || '0', 10);
    const fillId = parseInt((tag.match(/fillId="(\d+)"/) || [])[1] || '0', 10);
    const borderId = parseInt((tag.match(/borderId="(\d+)"/) || [])[1] || '0', 10);
    const applyFont = /applyFont="1"/.test(tag);
    const applyFill = /applyFill="1"/.test(tag);
    const applyBorder = /applyBorder="1"/.test(tag);
    const applyAlign = /applyAlignment="1"/.test(tag);
    const fmt = numFmtMap.get(numFmtId) || builtin.get(numFmtId);
    const font = fonts[fontId] || {};
    const fill = fills[fillId]?.fill;
    
  const style: CellStyle & { numberFormat?: string } = {};
    if (applyFont) {
      if (font.name) style.fontFamily = font.name;
      if (font.size) style.fontSize = font.size;
      if (font.bold) style.bold = font.bold;
      if (font.italic) style.italic = font.italic;
      if (font.underline) style.underline = font.underline;
      if (font.color) style.color = font.color;
      // Phase 1 UI: strikethrough, superscript, subscript
      if (font.strikethrough) (style as any).strikethrough = font.strikethrough;
      if (font.superscript) (style as any).superscript = font.superscript;
      if (font.subscript) (style as any).subscript = font.subscript;
    }
    // Apply fill if applyFill is set OR if fillId > 1 (Excel often omits applyFill but still uses fills)
    if ((applyFill || fillId > 1) && fill) style.fill = fill;
    if (applyBorder) {
      const b = borders[borderId];
      if (b) style.border = { ...b } as any;
    }
    if (applyAlign) {
      // Extract horizontal/vertical alignment and wrapping/rotation
      const start = xml.indexOf(tag);
      const after = xml.slice(start, start + 500); // small window
      const alignMatch = after.match(/<alignment[^>]*>/);
      if (alignMatch) {
        const t = alignMatch[0];
        const h = (t.match(/horizontal="(\w+)"/) || [])[1];
        const v = (t.match(/vertical="(\w+)"/) || [])[1];
        const wrap = (t.match(/wrapText="(\d+)"/) || [])[1];
        const rot = (t.match(/textRotation="(-?\d+)"/) || [])[1];
        const indent = (t.match(/indent="(\d+)"/) || [])[1]; // Phase 1 UI
        if (h === 'left' || h === 'center' || h === 'right') (style as any).align = h;
        if (v === 'top' || v === 'center' || v === 'bottom') (style as any).valign = v === 'center' ? 'middle' : (v as any);
        if (wrap) (style as any).wrap = wrap === '1';
        if (rot) (style as any).rotation = parseInt(rot, 10);
        if (indent) (style as any).indent = parseInt(indent, 10); // Phase 1 UI
      }
    }
    (style as any).textOverflow = 'clip';
    if (fmt) style.numberFormat = fmt;
    xfs.push(style);
  }
  
  return { xfStyles: xfs };
}

function argbToCss(argb: string): string {
  // ARGB hex (AARRGGBB) -> rgba or #RRGGBB (ignoring alpha if 0xFF)
  if (argb.length !== 8) return '#000000';
  const a = parseInt(argb.slice(0, 2), 16) / 255;
  const r = parseInt(argb.slice(2, 4), 16);
  const g = parseInt(argb.slice(4, 6), 16);
  const b = parseInt(argb.slice(6, 8), 16);
  if (a >= 0.999) return `#${argb.slice(2)}`;
  return `rgba(${r},${g},${b},${a.toFixed(3)})`;
}

// ----- Theme parsing -----
type ThemePalette = { colors: string[] }; // colors in #RRGGBB, index mapping per Excel: [lt1, dk1, lt2, dk2, accent1..accent6, hlink, folHlink]

function parseTheme(xml: string): ThemePalette | undefined {
  if (!xml) return undefined;
  // Extract clrScheme block
  const schemeMatch = xml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
  if (!schemeMatch) return undefined;
  const scheme = schemeMatch[0];
  // Roles in Excel theme index order
  const roles = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const colors: string[] = [];
  for (const role of roles) {
    const roleMatch = scheme.match(new RegExp(`<a:${role}[^>]*>([\\s\\S]*?)<\\/a:${role}>`));
    if (!roleMatch) { colors.push('#000000'); continue; }
    const roleXml = roleMatch[0];
    // Prefer explicit srgbClr
    const srgb = roleXml.match(/<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/);
    if (srgb) { colors.push(`#${(srgb[1] as string).toUpperCase()}`); continue; }
    // Optionally map schemeClr via val attr, but most themes include srgbClr; default to black
    colors.push('#000000');
  }
  return { colors };
}

function themeColorToCss(theme: ThemePalette, themeIdx: number, tint: number): string {
  // Excel theme color index mapping (approx): 0.. is usually [lt1, dk1, lt2, dk2, accent1..accent6, hyperlink, followedHyperlink]
  const base = theme.colors[themeIdx] || '#000000';
  const rgb = hexToRgb(base);
  const shaded = applyTint(rgb, tint);
  return rgbToHex(shaded);
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const h = hex.replace('#', '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return { r, g, b };
}

function rgbToHex({ r, g, b }: { r: number; g: number; b: number }): string {
  const to = (n: number) => n.toString(16).padStart(2, '0');
  return `#${to(clamp255(r))}${to(clamp255(g))}${to(clamp255(b))}`.toUpperCase();
}

function clamp255(n: number) { return Math.max(0, Math.min(255, Math.round(n))); }

// Excel tint/shade formula: https://stackoverflow.com/a/40520243
function applyTint(rgb: { r: number; g: number; b: number }, tint: number): { r: number; g: number; b: number } {
  // tint in [-1, 1]; positive = lighten, negative = darken
  const apply = (c: number) => {
    if (tint < 0) return c * (1 + tint);
    return c + (255 - c) * tint;
  };
  return { r: apply(rgb.r), g: apply(rgb.g), b: apply(rgb.b) } as any;
}
