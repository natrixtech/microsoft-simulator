export const FONTS: string[] = [
  'Arial',
  'Segoe UI',
  'Times New Roman',
  'Courier New',
  'Georgia',
  'Verdana',
  'Comic Sans MS',
];

export const COLORS: Record<string, string> = {
  black: '#000000',
  darkGray: '#434343',
  gray: '#999999',
  white: '#ffffff',
  red: '#cc0000',
  orange: '#e69138',
  yellow: '#f1c232',
  green: '#6aa84f',
  cyan: '#45818e',
  blue: '#3d85c6',
  purple: '#674ea7',
  magenta: '#a64d79',
  lightRed: '#f4cccc',
  lightOrange: '#fce5cd',
  lightYellow: '#fff2cc',
  lightGreen: '#d9ead3',
  lightCyan: '#d0e0e3',
  lightBlue: '#cfe2f3',
  lightPurple: '#d9d2e9',
  lightMagenta: '#ead1dc',
};

export function clone<T>(val: T): T {
  if (typeof structuredClone === 'function') {
    return structuredClone(val);
  }
  return JSON.parse(JSON.stringify(val));
}

export function esc(str: unknown): string {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

export function safeAsset(baseUrl: string, file: string): string {
  if (!file) return '';
  if (/^https?:\/\//i.test(file) || file.startsWith('data:')) return file;
  const base = (baseUrl || '').replace(/\/$/, '');
  const path = file.replace(/^\//, '');
  return base ? `${base}/${path}` : path;
}

export function matches(target: any, source: any, keys: string[]): boolean {
  if (!target || !source) return false;
  return keys.every((k) => source[k] === undefined || target[k] === source[k]);
}

export function sanitizeHtml(html: string, assetBaseUrl: string = ''): string {
  if (!html) return '';
  const div = document.createElement('div');
  div.innerHTML = html;

  div.querySelectorAll('script, style, iframe, object, embed, link').forEach((el) => el.remove());

  div.querySelectorAll('*').forEach((el) => {
    Array.from(el.attributes).forEach((attr) => {
      if (attr.name.startsWith('on') || attr.value.trim().toLowerCase().startsWith('javascript:')) {
        el.removeAttribute(attr.name);
      }
    });
  });

  div.querySelectorAll('img').forEach((img) => {
    const src = img.getAttribute('src') || '';
    if (src && !src.startsWith('data:') && !src.startsWith('http://') && !src.startsWith('https://')) {
      img.src = safeAsset(assetBaseUrl, src);
    }
  });

  return div.innerHTML;
}
