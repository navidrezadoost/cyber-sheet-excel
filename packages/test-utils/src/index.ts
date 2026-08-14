export function generateMatrix(rows: number, cols: number, fn: (r: number, c: number) => any = (r,c)=>`${r},${c}`) {
  const data: any[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: any[] = [];
    for (let c = 0; c < cols; c++) row.push(fn(r, c));
    data.push(row);
  }
  return data;
}

export function mockCanvas() {
  // Minimal mock for Canvas to avoid crash in JSDOM
  if (!(globalThis as any).HTMLCanvasElement) {
    (globalThis as any).HTMLCanvasElement = function HTMLCanvasElement() {} as any;
  }
  const proto: any = (globalThis as any).HTMLCanvasElement.prototype;
  // Always override getContext to avoid jsdom's not-implemented error
  Object.defineProperty(proto, 'getContext', {
    value: () => {
      const ctx: any = {
        save: () => {},
        restore: () => {},
        clearRect: () => {},
        fillRect: () => {},
        strokeRect: () => {},
        beginPath: () => {},
        closePath: () => {},
        moveTo: () => {},
        lineTo: () => {},
        arc: () => {},
        rect: () => {},
        clip: () => {},
        fill: () => {},
        stroke: () => {},
        fillText: () => {},
        strokeText: () => {},
        measureText: (t: string) => ({ width: t.length * 7 }),
        drawImage: () => {},
        createLinearGradient: () => ({
          addColorStop: () => {},
        }),
        translate: () => {},
        scale: () => {},
        rotate: () => {},
        setLineDash: () => {},
        setTransform: () => {},
        resetTransform: () => {},
        // common properties used by renderer
        font: '',
        fillStyle: '',
        strokeStyle: '',
        lineWidth: 1,
        textAlign: 'left',
        textBaseline: 'alphabetic',
      };
      return ctx;
    },
    configurable: true,
    writable: true,
  });
}

export function wait(ms: number) {
  return new Promise(res => setTimeout(res, ms));
}
