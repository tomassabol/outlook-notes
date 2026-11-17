import type { EditorFocus } from '../shared/types';
import { state } from './state';

/**
 * The editor runs as an extension page in an iframe. Keystrokes inside it never reach
 * Outlook, so typing "r" or pressing Delete can't trigger Outlook's keyboard shortcuts.
 */
let frame: HTMLIFrameElement | null = null;
let returnFocus: Element | null = null;

export function openEditor(convId: string, subject: string, focus: EditorFocus): void {
  frame?.remove();
  returnFocus = document.activeElement;
  const url = new URL(chrome.runtime.getURL('editor.html'));
  url.searchParams.set('conv', convId);
  url.searchParams.set('subject', subject.slice(0, 200));
  url.searchParams.set('focus', focus);
  url.searchParams.set('theme', state.theme);

  frame = document.createElement('iframe');
  frame.setAttribute('data-onl', '');
  frame.title = 'Note and labels';
  frame.src = url.href;
  Object.assign(frame.style, {
    position: 'fixed',
    inset: '0',
    width: '100vw',
    height: '100vh',
    border: '0',
    margin: '0',
    zIndex: '2147483647',
    background: 'transparent',
    colorScheme: 'normal',
  });
  const f = frame;
  f.addEventListener('load', () => f.focus());
  document.body.append(f);
}

export function initEditorFrame(): void {
  window.addEventListener('message', (e) => {
    if (!frame || e.source !== frame.contentWindow) return;
    if ((e.data as { type?: string })?.type === 'onl-close') {
      frame.remove();
      frame = null;
      (returnFocus as HTMLElement | null)?.focus?.();
    }
  });
}
