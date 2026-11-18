import tokensCss from '../shared/tokens.css?inline';
import { paletteTokenNames } from '../shared/palette';
import type { EditorFocus } from '../shared/types';
import { state } from './state';

/** Every Outlook (Fluent) token our styles reference, plus the category palette. */
const TOKEN_NAMES = [
  ...new Set([...tokensCss.matchAll(/var\((--[A-Za-z0-9]+)/g)].map((m) => m[1]).filter((n) => !/^--(onl|pill)-/.test(n))),
  ...paletteTokenNames(),
  '--durationFaster',
  '--colorTransparentStroke',
];

function outlookTokens(): Record<string, string> {
  const cs = getComputedStyle(document.documentElement);
  const out: Record<string, string> = {};
  for (const n of TOKEN_NAMES) {
    const v = cs.getPropertyValue(n).trim();
    if (v) out[n] = v;
  }
  return out;
}

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
  document.body.append(frame);
  // Focus the frame right away; the editor focuses its own field once shown. (Focusing the frame
  // again on 'load' would reset that.)
  frame.focus();
}

export function initEditorFrame(): void {
  window.addEventListener('message', (e) => {
    if (!frame || e.source !== frame.contentWindow) return;
    const type = (e.data as { type?: string })?.type;
    if (type === 'onl-ready') {
      frame.contentWindow!.postMessage({ type: 'onl-tokens', tokens: outlookTokens() }, new URL(chrome.runtime.getURL('')).origin);
    } else if (type === 'onl-close') {
      frame.remove();
      frame = null;
      (returnFocus as HTMLElement | null)?.focus?.();
    }
  });
}
