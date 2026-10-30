export interface PopupPosition {
  top: number;
  left: number;
}

export interface PopupPositionOptions {
  margin?: number;
  minMargin?: number;
}

/**
 * Computes fixed viewport coordinates (top, left) for a popup relative to its anchor element.
 * Left-aligned to button, top = button bottom + 2px.
 * Keeps popup inside viewport (shifts left if right overflow; flips above or shifts up if bottom overflow).
 */
export function computePopupPosition(
  anchor: HTMLElement | DOMRect,
  popup?: HTMLElement | { width: number; height: number } | null,
  options?: PopupPositionOptions
): PopupPosition {
  const margin = options?.margin ?? 2;
  const minMargin = options?.minMargin ?? 8;

  const rect = anchor instanceof HTMLElement ? anchor.getBoundingClientRect() : anchor;
  let left = rect.left;
  let top = rect.bottom + margin;

  const vpW = window.innerWidth || document.documentElement.clientWidth || 960;
  const vpH = window.innerHeight || document.documentElement.clientHeight || 540;

  let width = 0;
  let height = 0;
  if (popup instanceof HTMLElement) {
    const pRect = popup.getBoundingClientRect();
    width = pRect.width || popup.offsetWidth;
    height = pRect.height || popup.offsetHeight;
  } else if (popup) {
    width = popup.width;
    height = popup.height;
  }

  // Prevent right overflow: shift left
  if (width > 0 && left + width > vpW - minMargin) {
    left = Math.max(minMargin, vpW - width - minMargin);
  } else if (left < minMargin) {
    left = minMargin;
  }

  // Prevent bottom overflow: flip above anchor if room, or shift up
  if (height > 0 && top + height > vpH - minMargin) {
    if (rect.top - height - margin >= minMargin) {
      top = rect.top - height - margin;
    } else {
      top = Math.max(minMargin, vpH - height - minMargin);
    }
  }

  return {
    left: Math.round(left),
    top: Math.round(top),
  };
}

