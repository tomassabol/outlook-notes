import tokensCss from '../shared/tokens.css?inline';
import { icon } from '../shared/icons';
import css from './ui.css?inline';
import { state } from './state';
import type { Label } from '../shared/types';

const STYLE = tokensCss + '\n' + css;

/** Creates an element with an open shadow root holding our styles, isolated from Outlook's CSS. */
export function makeHost(attr: string, tag: 'span' | 'div' = 'span'): { host: HTMLElement; root: ShadowRoot } {
  const host = document.createElement(tag);
  host.setAttribute(attr, '');
  host.setAttribute('data-onl', '');
  host.dataset.theme = state.theme;
  const root = host.attachShadow({ mode: 'open' });
  const style = document.createElement('style');
  style.textContent = STYLE;
  root.append(style);
  return { host, root };
}

/** Replaces everything in a shadow root except its <style>. */
export function setContent(root: ShadowRoot, ...nodes: Node[]): void {
  const style = root.querySelector('style')!;
  root.replaceChildren(style, ...nodes);
}

let toastTimer: number | undefined;
export function toast(text: string): void {
  document.querySelector('[data-onl-toast]')?.remove();
  const { host, root } = makeHost('data-onl-toast', 'div');
  const box = document.createElement('div');
  box.className = 'toast';
  box.setAttribute('role', 'status');
  const span = document.createElement('span');
  span.textContent = text;
  box.append(icon('info16'), span);
  root.append(box);
  document.body.append(host);
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => host.remove(), 3500);
}

/** Copies text to the clipboard (Clipboard API, with an execCommand fallback). */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;top:-1000px;opacity:0';
    document.body.append(ta);
    ta.select();
    const ok = document.execCommand('copy');
    ta.remove();
    return ok;
  }
}

/** Makes an element copy `text` on click or Enter/Space. */
export function copyable(el: HTMLElement, text: string): HTMLElement {
  el.classList.add('copyable');
  el.setAttribute('role', 'button');
  el.tabIndex = 0;
  el.title = `Copy “${text}”`;
  const run = async (e: Event) => {
    e.preventDefault();
    e.stopPropagation();
    const ok = await copyText(text);
    toast(ok ? `Copied “${text}”` : 'Couldn’t copy to the clipboard');
    if (ok) {
      el.classList.add('copied');
      window.setTimeout(() => el.classList.remove('copied'), 700);
    }
  };
  el.addEventListener('click', run);
  el.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') void run(e);
  });
  return el;
}

/** Label pill copies its text without the emoji. */
export const labelCopyText = (l: Label) => l.text.trim() || l.emoji;
