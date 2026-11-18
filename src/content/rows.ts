import { hasContent } from '../shared/store';
import { h, pill } from '../shared/dom';
import { icon } from '../shared/icons';
import { applyPaletteVars } from '../shared/palette';
import type { ConvData, Label } from '../shared/types';
import { findRows, findSubjectEl } from './selectors';
import { dataFor, labelsOf, signature, state } from './state';
import { makeHost, setContent } from './ui';

const ATTR = 'data-onl-row';
const MAX_FULL = 3;
/** Share of the row width the tags may take; the subject truncates first, like Outlook categories. */
const SHARE = 0.36;
const MIN_BUDGET = 44;

/** Small emoji-only chip in the label's color (used when full tags don't fit). */
function chip(l: Label): HTMLSpanElement {
  const text = `${l.emoji} ${l.text}`.trim();
  const el = h('span', { class: 'chip', title: text, 'aria-label': text }, l.emoji || l.text.slice(0, 1).toUpperCase());
  applyPaletteVars(el, l.color);
  return el;
}

const more = (n: number, labels: Label[]) =>
  h('span', { class: 'more', title: labels.map((l) => `${l.emoji} ${l.text}`.trim()).join(', ') }, `+${n}`);

/**
 * Progressively more compact layouts. The first one that fits the budget is used, so a tag is
 * never cut in half.
 */
function layouts(note: boolean, labels: Label[]): (() => HTMLElement)[] {
  const box = (...kids: (Node | null)[]) =>
    h('span', { class: 'row' }, note ? h('span', { class: 'noteicon', title: 'Has a note' }, icon('note16')) : null, ...kids);
  const out: (() => HTMLElement)[] = [];
  if (labels.length) {
    // 1. full tags (up to 3) + "+N"
    out.push(() => box(...labels.slice(0, MAX_FULL).map((l) => pill(l, true)), labels.length > MAX_FULL ? more(labels.length - MAX_FULL, labels.slice(MAX_FULL)) : null));
    // 2. first tag full, the rest as chips
    if (labels.length > 1) out.push(() => box(pill(labels[0], true), ...labels.slice(1).map(chip)));
    // 3. all chips
    out.push(() => box(...labels.map(chip)));
    // 4..n. fewer chips + "+N"
    for (let k = labels.length - 1; k >= 1; k--) out.push(() => box(...labels.slice(0, k).map(chip), more(labels.length - k, labels.slice(k))));
    // last resort: just the count
    out.push(() => box(more(labels.length, labels)));
  } else {
    out.push(() => box());
  }
  return out;
}

function render(host: HTMLElement, d: ConvData, budget: number): void {
  const root = host.shadowRoot!;
  const options = layouts(!!d.n?.trim(), labelsOf(d));
  for (let i = 0; i < options.length; i++) {
    setContent(root, options[i]());
    const used = (root.lastElementChild as HTMLElement).getBoundingClientRect().width;
    if (used <= budget || i === options.length - 1) return;
  }
}

/**
 * Space the tags may use. Measured against the whole row (not the subject line, which may size
 * to its content and so depend on our own width).
 */
function budgetFor(row: HTMLElement): number {
  return Math.max(MIN_BUDGET, Math.floor(row.getBoundingClientRect().width * SHARE));
}

const observed = new WeakSet<Element>();
let resizeObserver: ResizeObserver | null = null;

/** Re-fit when the message list changes width (window resize, reading pane drag). */
function watchListWidth(row: HTMLElement, onResize: () => void): void {
  const list = row.closest('[role="listbox"]') ?? row.parentElement;
  if (!list || observed.has(list)) return;
  resizeObserver ??= new ResizeObserver(onResize);
  resizeObserver.observe(list);
  observed.add(list);
}

/** Adds/updates/removes the note icon + label tags on every rendered row. Idempotent. */
export function decorateRows(onResize: () => void): void {
  for (const { row, convId } of findRows()) {
    const d = dataFor(convId);
    let host = row.querySelector<HTMLElement>(`[${ATTR}]`);
    if (!hasContent(d)) {
      host?.remove();
      continue;
    }
    watchListWidth(row, onResize);

    if (!host) {
      host = makeHost(ATTR).host;
      const subject = findSubjectEl(row);
      if (subject) subject.after(host);
      else (row.querySelector('[data-convid]') ?? row).append(host);
    }
    const budget = budgetFor(row);
    // Width is bucketed so small layout jitters don't cause re-renders.
    const sig = signature(convId, d, String(Math.round(budget / 8)));
    if (host.dataset.sig === sig) continue;
    host.dataset.sig = sig;
    host.dataset.theme = state.theme;
    render(host, d, budget);
  }
}
