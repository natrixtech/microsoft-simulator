export const SLIDE_WIDTH = 960;
export const SLIDE_HEIGHT = 540;
export const SLIDE_ASPECT_RATIO = SLIDE_WIDTH / SLIDE_HEIGHT; // 16:9 = 1.777777...

export interface SlideViewportFit {
  scale: number;
  width: number;
  height: number;
  offsetX: number;
  offsetY: number;
  percent: number;
}

export function computeSlideFit(
  containerWidth: number,
  containerHeight: number,
  padding: number = 24
): SlideViewportFit {
  const availW = Math.max(100, containerWidth - padding * 2);
  const availH = Math.max(100, containerHeight - padding * 2);

  const scaleW = availW / SLIDE_WIDTH;
  const scaleH = availH / SLIDE_HEIGHT;
  const scale = Math.max(0.1, Math.min(scaleW, scaleH));

  const scaledW = Math.round(SLIDE_WIDTH * scale);
  const scaledH = Math.round(SLIDE_HEIGHT * scale);

  const offsetX = Math.max(0, Math.round((containerWidth - scaledW) / 2));
  const offsetY = Math.max(0, Math.round((containerHeight - scaledH) / 2));

  return {
    scale,
    width: scaledW,
    height: scaledH,
    offsetX,
    offsetY,
    percent: Math.round(scale * 100),
  };
}

export function screenToSlideCoords(
  screenX: number,
  screenY: number,
  scale: number,
  offsetX: number,
  offsetY: number
): { x: number; y: number } {
  return {
    x: Math.round((screenX - offsetX) / scale),
    y: Math.round((screenY - offsetY) / scale),
  };
}

export function slideToScreenCoords(
  slideX: number,
  slideY: number,
  scale: number,
  offsetX: number,
  offsetY: number
): { x: number; y: number } {
  return {
    x: Math.round(slideX * scale + offsetX),
    y: Math.round(slideY * scale + offsetY),
  };
}

