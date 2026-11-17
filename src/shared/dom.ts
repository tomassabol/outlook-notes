import { textColorFor } from './store';
import type { Label } from './types';

type Child = Node | string | null | undefined | false;
type Props = Record<string, string | boolean | number | EventListener | undefined>;

/** Tiny element builder. `on*` props become listeners; `class` sets className; text is always set via textContent. */
export function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Props = {}, ...children: Child[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v === undefined || v === false) continue;
    if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'class') el.className = String(v);
    else if (k in el && typeof v !== 'string') (el as unknown as Record<string, unknown>)[k] = v;
    else el.setAttribute(k, v === true ? '' : String(v));
  }
  for (const c of children) if (c !== null && c !== undefined && c !== false) el.append(c);
  return el;
}

export function pill(label: Label): HTMLSpanElement {
  const text = `${label.emoji} ${label.text}`.trim();
  const el = h(
    'span',
    { class: 'pill', title: text },
    label.emoji ? h('span', { class: 'emoji' }, label.emoji) : null,
    label.text ? h('span', { class: 'txt' }, label.text) : null,
  );
  el.style.background = label.color;
  el.style.color = textColorFor(label.color);
  return el;
}

export function injectStyle(root: Document | ShadowRoot, css: string): HTMLStyleElement {
  const s = document.createElement('style');
  s.textContent = css;
  (root instanceof Document ? root.head : root).append(s);
  return s;
}
