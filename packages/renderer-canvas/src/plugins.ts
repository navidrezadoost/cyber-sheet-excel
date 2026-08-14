import type { Address, CellStyle, ExtendedCellValue } from '@cyber-sheet/core';

/**
 * Plugin system for extending renderer capabilities
 */

export type ColorTransformFn = (color: string, context: {
  addr: Address;
  value: ExtendedCellValue;
  style?: CellStyle;
}) => string;

export type CellColorFn = (context: {
  addr: Address;
  value: ExtendedCellValue;
  style?: CellStyle;
  min?: number;
  max?: number;
}) => string | undefined;

type RenderCellContext = {
  addr: Address;
  value: ExtendedCellValue;
  style?: CellStyle;
};

export type AccessibilityMode = 'none' | 'high-contrast' | 'deuteranopia' | 'protanopia' | 'tritanopia';

export interface RenderPlugin {
  id: string;
  priority?: number; // higher runs later
  
  // Color grading: transform any color before rendering
  transformColor?(color: string, context: RenderCellContext): string;
  
  // Heatmap: compute background color based on cell value
  getCellBackground?(context: RenderCellContext & { min?: number; max?: number }): string | undefined;
  
  // Custom font: override font string
  transformFont?(font: string, context: RenderCellContext): string;
  
  // Post-render hook for overlays
  afterCellRender?(ctx: CanvasRenderingContext2D, rect: { x: number; y: number; w: number; h: number }, context: RenderCellContext): void;
}

/**
 * Color grading plugin with HSL adjustments
 */
export class ColorGradingPlugin implements RenderPlugin {
  id = 'color-grading';
  priority = 10;
  
  constructor(
    private hue = 0,        // -180 to 180
    private saturation = 0, // -100 to 100
    private lightness = 0   // -100 to 100
  ) {}
  
  setHSL(h: number, s: number, l: number) {
    this.hue = h;
    this.saturation = s;
    this.lightness = l;
  }
  
  transformColor(color: string): string {
    if (!color || color === 'transparent') return color;
    const hsl = this.rgbToHsl(color);
    if (!hsl) return color;
    
    let [h, s, l] = hsl;
    h = (h + this.hue + 360) % 360;
    s = Math.max(0, Math.min(100, s + this.saturation));
    l = Math.max(0, Math.min(100, l + this.lightness));
    
    return this.hslToRgb(h, s, l);
  }
  
  private rgbToHsl(color: string): [number, number, number] | null {
    // Parse hex or rgb()
    let r = 0, g = 0, b = 0;
    if (color.startsWith('#')) {
      const hex = color.slice(1);
      if (hex.length === 6) {
        r = parseInt(hex.slice(0, 2), 16) / 255;
        g = parseInt(hex.slice(2, 4), 16) / 255;
        b = parseInt(hex.slice(4, 6), 16) / 255;
      } else if (hex.length === 3) {
        r = parseInt(hex[0] + hex[0], 16) / 255;
        g = parseInt(hex[1] + hex[1], 16) / 255;
        b = parseInt(hex[2] + hex[2], 16) / 255;
      }
    } else if (color.startsWith('rgb')) {
      const match = color.match(/\d+/g);
      if (match && match.length >= 3) {
        r = parseInt(match[0]) / 255;
        g = parseInt(match[1]) / 255;
        b = parseInt(match[2]) / 255;
      }
    } else {
      return null;
    }
    
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
        case g: h = ((b - r) / d + 2) / 6; break;
        case b: h = ((r - g) / d + 4) / 6; break;
      }
    }
    
    return [h * 360, s * 100, l * 100];
  }
  
  private hslToRgb(h: number, s: number, l: number): string {
    h /= 360;
    s /= 100;
    l /= 100;
    
    let r, g, b;
    if (s === 0) {
      r = g = b = l;
    } else {
      const hue2rgb = (p: number, q: number, t: number) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1/6) return p + (q - p) * 6 * t;
        if (t < 1/2) return q;
        if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
        return p;
      };
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;
      r = hue2rgb(p, q, h + 1/3);
      g = hue2rgb(p, q, h);
      b = hue2rgb(p, q, h - 1/3);
    }
    
    return `rgb(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)})`;
  }
}

/**
 * Accessibility plugin with colorblind filters and high contrast
 */
export class AccessibilityPlugin implements RenderPlugin {
  id = 'accessibility';
  priority = 5;
  
  constructor(private mode: AccessibilityMode = 'none') {}
  
  setMode(mode: AccessibilityMode) {
    this.mode = mode;
  }
  
  transformColor(color: string): string {
    if (this.mode === 'none' || !color) return color;
    
    if (this.mode === 'high-contrast') {
      return this.toHighContrast(color);
    }
    
    // Colorblind simulation matrices
    return this.applyColorblindFilter(color, this.mode);
  }
  
  private toHighContrast(color: string): string {
    const hsl = this.parseToHsl(color);
    if (!hsl) return color;
    const [h, s, l] = hsl;
    // Push to extremes
    const newL = l < 50 ? 0 : 100;
    return `hsl(${h}, ${s}%, ${newL}%)`;
  }
  
  private applyColorblindFilter(color: string, mode: AccessibilityMode): string {
    // Simplified colorblind simulation
    const rgb = this.parseToRgb(color);
    if (!rgb) return color;
    let [r, g, b] = rgb;
    
    switch (mode) {
      case 'deuteranopia': // green-blind
        r = 0.625 * r + 0.375 * g;
        g = 0.7 * g + 0.3 * r;
        break;
      case 'protanopia': // red-blind
        r = 0.567 * r + 0.433 * g;
        g = 0.558 * g + 0.442 * r;
        break;
      case 'tritanopia': // blue-blind
        b = 0.95 * b + 0.05 * g;
        g = 0.433 * g + 0.567 * b;
        break;
    }
    
    return `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
  }
  
  private parseToHsl(color: string): [number, number, number] | null {
    const rgb = this.parseToRgb(color);
    if (!rgb) return null;
    const [r, g, b] = rgb.map(x => x / 255);
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
        case g: h = ((b - r) / d + 2) / 6; break;
        case b: h = ((r - g) / d + 4) / 6; break;
      }
    }
    return [h * 360, s * 100, l * 100];
  }
  
  private parseToRgb(color: string): [number, number, number] | null {
    if (color.startsWith('#')) {
      const hex = color.slice(1);
      if (hex.length === 6) {
        return [
          parseInt(hex.slice(0, 2), 16),
          parseInt(hex.slice(2, 4), 16),
          parseInt(hex.slice(4, 6), 16),
        ];
      } else if (hex.length === 3) {
        return [
          parseInt(hex[0] + hex[0], 16),
          parseInt(hex[1] + hex[1], 16),
          parseInt(hex[2] + hex[2], 16),
        ];
      }
    } else if (color.startsWith('rgb')) {
      const match = color.match(/\d+/g);
      if (match && match.length >= 3) {
        return [parseInt(match[0]), parseInt(match[1]), parseInt(match[2])];
      }
    }
    return null;
  }
}

/**
 * Heatmap plugin for value-based color gradients
 */
export class HeatmapPlugin implements RenderPlugin {
  id = 'heatmap';
  priority = 20;
  
  constructor(
    private enabled = true,
    private colorScale: 'red-green' | 'blue-red' | 'grayscale' = 'red-green'
  ) {}
  
  setEnabled(enabled: boolean) {
    this.enabled = enabled;
  }
  
  setColorScale(scale: 'red-green' | 'blue-red' | 'grayscale') {
    this.colorScale = scale;
  }
  
  getCellBackground(context: RenderCellContext & { min?: number; max?: number }): string | undefined {
    if (!this.enabled) return undefined;
    const { value, min = 0, max = 100 } = context;
    if (typeof value !== 'number') return undefined;
    
    const range = max - min;
    if (range === 0) return undefined;
    
    const normalized = Math.max(0, Math.min(1, (value - min) / range));
    return this.getColor(normalized);
  }
  
  private getColor(t: number): string {
    switch (this.colorScale) {
      case 'red-green':
        // 0 = red, 0.5 = yellow, 1 = green
        if (t < 0.5) {
          const r = 255;
          const g = Math.round(t * 2 * 255);
          return `rgb(${r}, ${g}, 0)`;
        } else {
          const r = Math.round((1 - t) * 2 * 255);
          const g = 255;
          return `rgb(${r}, ${g}, 0)`;
        }
      case 'blue-red':
        // 0 = blue, 1 = red
        const r = Math.round(t * 255);
        const b = Math.round((1 - t) * 255);
        return `rgb(${r}, 0, ${b})`;
      case 'grayscale':
        const gray = Math.round(t * 255);
        return `rgb(${gray}, ${gray}, ${gray})`;
      default:
        return 'transparent';
    }
  }
}

/**
 * Custom font plugin
 */
export class CustomFontPlugin implements RenderPlugin {
  id = 'custom-font';
  priority = 15;
  
  private fontMap = new Map<string, string>();
  
  registerFont(originalFamily: string, customFamily: string) {
    this.fontMap.set(originalFamily.toLowerCase(), customFamily);
  }
  
  transformFont(font: string): string {
    // Parse font string: "italic bold 14px Arial, sans-serif"
    const parts = font.split(' ');
    const familyIndex = parts.findIndex(p => p.includes('px')) + 1;
    if (familyIndex > 0 && familyIndex < parts.length) {
      const families = parts.slice(familyIndex).join(' ').split(',');
      const mapped = families.map(f => {
        const trimmed = f.trim().toLowerCase();
        return this.fontMap.get(trimmed) || f.trim();
      });
      return parts.slice(0, familyIndex).join(' ') + ' ' + mapped.join(', ');
    }
    return font;
  }
  
  async loadFont(family: string, url: string): Promise<void> {
    if (typeof FontFace === 'undefined') {
      console.warn('FontFace API not available');
      return;
    }
    const fontFace = new FontFace(family, `url(${url})`);
    await fontFace.load();
    (document as any).fonts.add(fontFace);
  }
}
