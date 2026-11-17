import baseCss from '../shared/base.css?inline';
import css from './ui.css?inline';
import { state } from './state';

const STYLE = baseCss + '\n' + css;

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
  box.textContent = text;
  root.append(box);
  document.body.append(host);
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => host.remove(), 3500);
}
