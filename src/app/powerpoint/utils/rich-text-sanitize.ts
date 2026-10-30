/**
 * Sanitizer for rich text in PowerPoint slides.
 * Allows only: <div>, <br>, <ul>, <ol>, <li>, <b>, <i>, <u>, <span style>
 * with font-family, font-size (pt/px), color, text-decoration, text-align, line-height.
 */

export function sanitizeRichText(html: string): string {
  if (!html || typeof html !== 'string') return '';
  if (typeof DOMParser === 'undefined') {
    return html;
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  cleanNode(doc.body);
  return doc.body.innerHTML;
}

export function plainTextToHtml(text: string, isBullet: boolean = false): string {
  if (!text) return '';
  const lines = text.split('\n');
  if (isBullet) {
    const lis = lines.map((l) => `<li>${escapeHtml(l) || '<br>'}</li>`).join('');
    return `<ul>${lis}</ul>`;
  }
  return lines.map((l) => `<div>${escapeHtml(l) || '<br>'}</div>`).join('');
}

export function htmlToPlainText(html: string): string {
  if (!html) return '';
  if (typeof DOMParser === 'undefined') return html.replace(/<[^>]+>/g, '');
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return (doc.body.innerText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function cleanNode(node: Node): void {
  const childNodes = Array.from(node.childNodes);
  for (const child of childNodes) {
    if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as HTMLElement;
      const rawTag = el.tagName.toUpperCase();

      if (['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED'].includes(rawTag)) {
        el.remove();
        continue;
      }

      // Normalize synonymous tags
      let tag = rawTag;
      if (rawTag === 'STRONG') tag = 'B';
      else if (rawTag === 'EM') tag = 'I';
      else if (rawTag === 'P') tag = 'DIV';

      const allowedTags = ['DIV', 'BR', 'UL', 'OL', 'LI', 'B', 'I', 'U', 'SPAN'];
      if (!allowedTags.includes(tag)) {
        // Unwrap disallowed tags
        cleanNode(el);
        while (el.firstChild) {
          el.parentNode?.insertBefore(el.firstChild, el);
        }
        el.remove();
        continue;
      }

      let currentEl = el;
      if (tag !== rawTag) {
        const replacement = el.ownerDocument.createElement(tag.toLowerCase());
        while (el.firstChild) {
          replacement.appendChild(el.firstChild);
        }
        if (el.hasAttribute('style')) {
          replacement.setAttribute('style', el.getAttribute('style') || '');
        }
        el.parentNode?.replaceChild(replacement, el);
        currentEl = replacement;
      }

      // Allowed CSS style properties
      const allowedStyles = [
        'font-family',
        'font-size',
        'font-weight',
        'font-style',
        'color',
        'text-decoration',
        'text-align',
        'line-height',
      ];

      const computedStyles: string[] = [];
      for (const styleName of allowedStyles) {
        const camel = styleName.replace(/-([a-z])/g, (_, g) => g.toUpperCase());
        let val = (currentEl.style as any)[camel];
        if (!val && styleName === 'text-decoration') {
          val = (currentEl.style as any).textDecorationLine;
        }
        if (val) {
          computedStyles.push(`${styleName}: ${val}`);
        }
      }

      // Remove all attributes except the cleaned style
      const attrNames = Array.from(currentEl.attributes).map((a) => a.name);
      for (const attr of attrNames) {
        currentEl.removeAttribute(attr);
      }

      if (computedStyles.length > 0) {
        currentEl.setAttribute('style', computedStyles.join('; '));
      }

      cleanNode(currentEl);
    }
  }
}

