/**
 * DrawingLayer.ts
 *
 * Kernel-level drawing layer for managing floating objects on top of the cell grid.
 * Supports pictures, shapes, form controls, charts, text boxes, slicers, and timelines.
 *
 * Key responsibilities:
 * - Object lifecycle (CRUD)
 * - Z-ordering (bring to front, send to back)
 * - Hit testing for mouse interactions
 * - Selection management
 * - Grouping/ungrouping objects
 * - Serialization for persistence and undo/redo
 */

// ─── Base Drawing Object Interface ─────────────────────────────────────────

export type ImagePlacement = 'cell' | 'floating';

export interface CellImageAnchor {
  row: number;
  col: number;
  endRow?: number;
  endCol?: number;
  offsetX?: number;
  offsetY?: number;
}

export interface DrawingObject {
  id: string;
  type: 'picture' | 'shape' | 'icon' | 'formControl' | 'chart' | 'textBox' | 'slicer' | 'timeline';
  name: string;
  position: { x: number; y: number }; // Pixels from cell A1 top-left
  size: { width: number; height: number }; // Display size in pixels
  rotation: number; // Degrees (0-360)
  zIndex: number;
  locked: boolean;
  visible: boolean;
  altText: string;
  anchor?: {
    moveWithCells: boolean;
    resizeWithCells: boolean;
    colOffset: number;
    rowOffset: number;
  };
}

// ─── Specific Object Types ─────────────────────────────────────────────────

export interface PictureObject extends DrawingObject {
  type: 'picture';
  source: string; // Data URI, blob URL, or file path
  sourceType: 'dataUri' | 'blob' | 'url' | 'stockImage' | 'svg';
  naturalWidth: number;
  naturalHeight: number;
  placement?: ImagePlacement;
  cellAnchor?: CellImageAnchor;
  clipToCell?: boolean;
  loadedImage?: HTMLImageElement; // Cached loaded image for rendering
  cropSettings?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

export type ShapeType =
  | 'line' | 'arrow' | 'elbowConnector' | 'curve' | 'freeform' | 'scribble'
  | 'rectangle' | 'roundedRectangle' | 'snipSingle' | 'snipSameCorner' | 'snipDiagonal'
  | 'oval' | 'triangle' | 'diamond' | 'parallelogram' | 'trapezoid'
  | 'pentagon' | 'hexagon' | 'heptagon' | 'octagon'
  | 'plus' | 'cross' | 'cube' | 'cylinder' | 'brace' | 'bracket'
  | 'sun' | 'moon' | 'cloud' | 'heart' | 'lightning' | 'smileyFace'
  | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow'
  | 'flowchartProcess' | 'flowchartDecision' | 'flowchartData'
  | 'star5' | 'star6' | 'banner' | 'callout';

export interface FillProperties {
  type: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture';
  color?: string;
  transparency?: number; // 0-100
}

export interface LineProperties {
  color: string;
  width: number; // pixels
  style: 'solid' | 'dashed' | 'dotted' | 'double' | 'none';
  dashType?: 'roundDot' | 'squareDot' | 'dash' | 'dashDot' | 'longDash';
}

export interface ShadowProperties {
  type: 'none' | 'outer' | 'inner' | 'perspective';
  color: string;
  blur: number;
  angle: number;
  distance: number;
  transparency: number;
}

export interface ShapeObject extends DrawingObject {
  type: 'shape';
  shapeType: ShapeType;
  fill: FillProperties;
  line: LineProperties;
  shadow?: ShadowProperties;
  text?: string; // Text inside the shape
  textStyle?: {
    fontFamily: string;
    fontSize: number;
    color: string;
    bold: boolean;
    italic: boolean;
  };
}

export interface IconObject extends DrawingObject {
  type: 'icon';
  iconId: string;
  svgContent: string;
  viewBox: string;
  fillColor: string;
  loadedImage?: HTMLImageElement;
}

export type FormControlType =
  | 'checkbox' | 'button' | 'comboBox' | 'listBox'
  | 'spinButton' | 'scrollBar' | 'optionButton' | 'groupBox' | 'label';

export interface FormControlProperties {
  enabled: boolean;
  printObject: boolean;
  checked?: boolean;
  label?: string;
  inputRange?: string; // Range reference (e.g., "A1:A10")
  selectedIndex?: number;
  dropDownLines?: number;
  minValue?: number;
  maxValue?: number;
  incrementalChange?: number;
  pageChange?: number;
  buttonText?: string;
  macroName?: string;
  threeDShading?: boolean;
  lockText?: boolean;
  mixed?: boolean;
  altTextTitle?: string;
  altTextDescription?: string;
}

export interface FormControlObject extends DrawingObject {
  type: 'formControl';
  controlType: FormControlType;
  linkedCell?: string; // Cell reference (e.g., "B5")
  controlProperties: FormControlProperties;
}

export interface WordArtStyle {
  presetId: string;
  gradientFrom: string;
  gradientTo: string;
  outlineColor?: string;
  outlineWidth?: number;
  shadow?: boolean;
}

export interface TextBoxObject extends DrawingObject {
  type: 'textBox';
  text: string;
  wordArtStyle?: WordArtStyle;
  textStyle: {
    fontFamily: string;
    fontSize: number;
    color: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    align: 'left' | 'center' | 'right';
    valign: 'top' | 'middle' | 'bottom';
  };
  fill: FillProperties;
  border: LineProperties;
}

export interface ChartObject extends DrawingObject {
  type: 'chart';
  chartType: 'column' | 'line' | 'pie' | 'bar' | 'area' | 'scatter' | 'other';
  dataRange: string; // Range reference
  chartData?: any; // Serialized chart configuration
}

// ─── Rect for hit testing ───────────────────────────────────────────────────

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

// ─── Event types ───────────────────────────────────────────────────────────

export type DrawingLayerEvent = 'changed' | 'selectionChanged' | 'objectAdded' | 'objectRemoved';

// ─── Simple event emitter ───────────────────────────────────────────────────

class EventEmitter {
  private listeners: Map<string, Set<(...args: any[]) => void>> = new Map();

  on(event: string, callback: (...args: any[]) => void): void {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, new Set());
    }
    this.listeners.get(event)!.add(callback);
  }

  off(event: string, callback: (...args: any[]) => void): void {
    this.listeners.get(event)?.delete(callback);
  }

  emit(event: string, ...args: any[]): void {
    this.listeners.get(event)?.forEach(cb => cb(...args));
  }
}

// ─── DrawingLayer ───────────────────────────────────────────────────────────

export interface SerializedDrawingLayer {
  objects: DrawingObject[];
  zOrder: string[];
}

export interface AddPictureOptions {
  id?: string;
  name?: string;
  source: string;
  sourceType?: PictureObject['sourceType'];
  naturalWidth?: number;
  naturalHeight?: number;
  position?: { x: number; y: number };
  size?: { width: number; height: number };
  placement?: ImagePlacement;
  cellAnchor?: CellImageAnchor;
  clipToCell?: boolean;
  altText?: string;
  locked?: boolean;
  visible?: boolean;
}

export class DrawingLayer {
  private objects: Map<string, DrawingObject> = new Map();
  private zOrder: string[] = []; // Object IDs in z-order (back to front)
  private selection: Set<string> = new Set();
  private eventEmitter: EventEmitter = new EventEmitter();
  private nextObjectId: number = 1;

  // ── CRUD ──────────────────────────────────────────────────────────────────

  addPicture(options: AddPictureOptions): PictureObject {
    const naturalWidth = options.naturalWidth ?? options.size?.width ?? 0;
    const naturalHeight = options.naturalHeight ?? options.size?.height ?? 0;
    const picture: PictureObject = {
      id: options.id ?? '',
      type: 'picture',
      name: options.name ?? 'Picture',
      source: options.source,
      sourceType: options.sourceType ?? 'url',
      naturalWidth,
      naturalHeight,
      position: options.position ?? { x: 0, y: 0 },
      size: options.size ?? {
        width: naturalWidth,
        height: naturalHeight,
      },
      rotation: 0,
      zIndex: this.zOrder.length + 1,
      locked: options.locked ?? false,
      visible: options.visible ?? true,
      altText: options.altText ?? '',
      placement: options.placement ?? (options.cellAnchor ? 'cell' : 'floating'),
      cellAnchor: options.cellAnchor ? { ...options.cellAnchor } : undefined,
      clipToCell: options.clipToCell ?? Boolean(options.cellAnchor),
      anchor: options.cellAnchor
        ? {
            moveWithCells: true,
            resizeWithCells: options.clipToCell ?? true,
            colOffset: options.cellAnchor.offsetX ?? 0,
            rowOffset: options.cellAnchor.offsetY ?? 0,
          }
        : undefined,
    };

    this.addObject(picture);
    return picture;
  }

  addObject(obj: DrawingObject): void {
    if (!obj.id) {
      obj.id = this.generateId();
    }
    this.objects.set(obj.id, obj);
    this.zOrder.push(obj.id);
    this.eventEmitter.emit('objectAdded', obj);
    this.eventEmitter.emit('changed');
  }

  removeObject(id: string): DrawingObject | undefined {
    const obj = this.objects.get(id);
    if (!obj) return undefined;

    this.objects.delete(id);
    this.zOrder = this.zOrder.filter(objId => objId !== id);
    this.selection.delete(id);

    this.eventEmitter.emit('objectRemoved', obj);
    this.eventEmitter.emit('changed');
    return obj;
  }

  getObject(id: string): DrawingObject | undefined {
    return this.objects.get(id);
  }

  getAllObjects(): DrawingObject[] {
    return this.zOrder.map(id => this.objects.get(id)!).filter(obj => obj !== undefined);
  }

  updateObject(id: string, updates: Partial<DrawingObject>): void {
    const obj = this.objects.get(id);
    if (!obj) return;

    Object.assign(obj, updates);
    this.eventEmitter.emit('changed');
  }

  // ── Hit Testing ────────────────────────────────────────────────────────────

  getObjectsInRect(rect: Rect): DrawingObject[] {
    return this.getAllObjects().filter(obj => {
      if (!obj.visible) return false;

      const objRect = this.getObjectRect(obj);
      return this.rectsIntersect(rect, objRect);
    });
  }

  getObjectAtPoint(x: number, y: number): DrawingObject | undefined {
    // Iterate from front to back (reverse z-order)
    for (let i = this.zOrder.length - 1; i >= 0; i--) {
      const obj = this.objects.get(this.zOrder[i]);
      if (!obj || !obj.visible) continue;

      if (this.pointInObject(x, y, obj)) {
        return obj;
      }
    }
    return undefined;
  }

  private pointInObject(x: number, y: number, obj: DrawingObject): boolean {
    const rect = this.getObjectRect(obj);
    return x >= rect.x && x <= rect.x + rect.width &&
           y >= rect.y && y <= rect.y + rect.height;
  }

  private getObjectRect(obj: DrawingObject): Rect {
    return {
      x: obj.position.x,
      y: obj.position.y,
      width: obj.size.width,
      height: obj.size.height,
    };
  }

  private rectsIntersect(r1: Rect, r2: Rect): boolean {
    return !(r1.x + r1.width < r2.x ||
             r2.x + r2.width < r1.x ||
             r1.y + r1.height < r2.y ||
             r2.y + r2.height < r1.y);
  }

  // ── Z-order ────────────────────────────────────────────────────────────────

  bringToFront(id: string): void {
    this.zOrder = this.zOrder.filter(objId => objId !== id);
    this.zOrder.push(id);
    this.eventEmitter.emit('changed');
  }

  sendToBack(id: string): void {
    this.zOrder = this.zOrder.filter(objId => objId !== id);
    this.zOrder.unshift(id);
    this.eventEmitter.emit('changed');
  }

  bringForward(id: string): void {
    const idx = this.zOrder.indexOf(id);
    if (idx >= 0 && idx < this.zOrder.length - 1) {
      [this.zOrder[idx], this.zOrder[idx + 1]] = [this.zOrder[idx + 1], this.zOrder[idx]];
      this.eventEmitter.emit('changed');
    }
  }

  sendBackward(id: string): void {
    const idx = this.zOrder.indexOf(id);
    if (idx > 0) {
      [this.zOrder[idx], this.zOrder[idx - 1]] = [this.zOrder[idx - 1], this.zOrder[idx]];
      this.eventEmitter.emit('changed');
    }
  }

  // ── Selection ──────────────────────────────────────────────────────────────

  selectObject(id: string, multiSelect: boolean = false): void {
    if (!multiSelect) {
      this.selection.clear();
    }
    this.selection.add(id);
    this.eventEmitter.emit('selectionChanged', Array.from(this.selection));
  }

  deselectObject(id: string): void {
    this.selection.delete(id);
    this.eventEmitter.emit('selectionChanged', Array.from(this.selection));
  }

  deselectAll(): void {
    this.selection.clear();
    this.eventEmitter.emit('selectionChanged', []);
  }

  getSelectedObjects(): DrawingObject[] {
    return Array.from(this.selection)
      .map(id => this.objects.get(id))
      .filter((obj): obj is DrawingObject => obj !== undefined);
  }

  getSelectedIds(): string[] {
    return Array.from(this.selection);
  }

  isSelected(id: string): boolean {
    return this.selection.has(id);
  }

  // ── Position/Size ──────────────────────────────────────────────────────────

  moveObject(id: string, delta: { x: number; y: number }): void {
    const obj = this.objects.get(id);
    if (!obj || obj.locked) return;

    obj.position.x += delta.x;
    obj.position.y += delta.y;
    this.eventEmitter.emit('changed');
  }

  setObjectPosition(id: string, position: { x: number; y: number }): void {
    const obj = this.objects.get(id);
    if (!obj || obj.locked) return;

    obj.position = position;
    this.eventEmitter.emit('changed');
  }

  resizeObject(id: string, newSize: { width: number; height: number }, anchor: 'topLeft' | 'center' = 'topLeft'): void {
    const obj = this.objects.get(id);
    if (!obj || obj.locked) return;

    if (anchor === 'center') {
      const deltaW = newSize.width - obj.size.width;
      const deltaH = newSize.height - obj.size.height;
      obj.position.x -= deltaW / 2;
      obj.position.y -= deltaH / 2;
    }

    obj.size = newSize;
    this.eventEmitter.emit('changed');
  }

  rotateObject(id: string, degrees: number): void {
    const obj = this.objects.get(id);
    if (!obj || obj.locked) return;

    obj.rotation = (obj.rotation + degrees) % 360;
    this.eventEmitter.emit('changed');
  }

  setObjectRotation(id: string, degrees: number): void {
    const obj = this.objects.get(id);
    if (!obj || obj.locked) return;

    obj.rotation = ((degrees % 360) + 360) % 360;
    this.eventEmitter.emit('changed');
  }


  // ── Grouping ───────────────────────────────────────────────────────────────

  groupObjects(ids: string[]): string {
    // TODO: Implement grouping logic
    const groupId = this.generateId();
    console.warn('groupObjects not fully implemented yet:', ids);
    return groupId;
  }

  ungroupObjects(groupId: string): string[] {
    // TODO: Implement ungrouping logic
    console.warn('ungroupObjects not fully implemented yet:', groupId);
    return [];
  }

  // ── Serialization ──────────────────────────────────────────────────────────

  serialize(): SerializedDrawingLayer {
    return {
      objects: this.getAllObjects(),
      zOrder: [...this.zOrder],
    };
  }

  deserialize(data: SerializedDrawingLayer): void {
    this.objects.clear();
    this.zOrder = [];
    this.selection.clear();

    for (const obj of data.objects) {
      this.objects.set(obj.id, obj);
    }
    this.zOrder = [...data.zOrder];
    this.nextObjectId = this.getNextGeneratedId();

    this.eventEmitter.emit('changed');
  }

  clear(): void {
    this.objects.clear();
    this.zOrder = [];
    this.selection.clear();
    this.eventEmitter.emit('changed');
  }

  // ── Events ─────────────────────────────────────────────────────────────────

  on(event: DrawingLayerEvent, callback: (...args: any[]) => void): void {
    this.eventEmitter.on(event, callback);
  }

  off(event: DrawingLayerEvent, callback: (...args: any[]) => void): void {
    this.eventEmitter.off(event, callback);
  }

  // ── Utilities ──────────────────────────────────────────────────────────────

  private generateId(): string {
    return `obj_${this.nextObjectId++}`;
  }

  private getNextGeneratedId(): number {
    let max = 0;
    for (const id of this.objects.keys()) {
      const match = /^obj_(\d+)$/.exec(id);
      if (match) max = Math.max(max, Number(match[1]));
    }
    return max + 1;
  }

  getObjectCount(): number {
    return this.objects.size;
  }

  getSelectedCount(): number {
    return this.selection.size;
  }
}
