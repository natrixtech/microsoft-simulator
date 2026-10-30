import { Pipe, PipeTransform } from '@angular/core';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { sanitizeRichText } from '../utils/rich-text-sanitize';

@Pipe({
  name: 'pptSafeHtml',
  standalone: true,
  pure: true,
})
export class PptSafeHtmlPipe implements PipeTransform {
  constructor(private sanitizer: DomSanitizer) {}

  transform(html: string | null | undefined): SafeHtml {
    return this.sanitizer.bypassSecurityTrustHtml(sanitizeRichText(html || ''));
  }
}

