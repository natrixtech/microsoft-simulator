export const PPT_FONTS: string[] = [
  'Arial',
  'Georgia',
  'Times New Roman',
  'Verdana',
  'Courier New',
  'Comic Sans MS',
];

export const PPT_FONT_SIZES: number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
];

export function getNextFontSize(currentSize: number, direction: 'increase' | 'decrease'): number {
  const sizes = PPT_FONT_SIZES;
  const cur = currentSize;
  if (direction === 'increase') {
    const next = sizes.find((s) => s > cur);
    return next !== undefined ? next : Math.min(96, cur + 4);
  } else {
    for (let i = sizes.length - 1; i >= 0; i--) {
      if (sizes[i] < cur) return sizes[i];
    }
    return Math.max(8, cur - 4);
  }
}
