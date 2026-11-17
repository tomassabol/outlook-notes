import { hasContent } from '../shared/store';
import { h, pill } from '../shared/dom';
import { findRows, findSubjectEl } from './selectors';
import { dataFor, labelsOf, signature, state } from './state';
import { makeHost, setContent } from './ui';

const ATTR = 'data-onl-row';
const MAX_PILLS = 3;

/** Adds/updates/removes the 📝 + label pills on every rendered row. Idempotent. */
export function decorateRows(): void {
  for (const { row, convId } of findRows()) {
    const d = dataFor(convId);
    let host = row.querySelector<HTMLElement>(`[${ATTR}]`);
    if (!hasContent(d)) {
      host?.remove();
      continue;
    }
    const sig = signature(convId, d);
    if (host?.dataset.sig === sig) continue;

    if (!host) {
      host = makeHost(ATTR).host;
      const subject = findSubjectEl(row);
      if (subject) subject.after(host);
      else (row.querySelector('[data-convid]') ?? row).append(host);
    }
    host.dataset.sig = sig;
    host.dataset.theme = state.theme;

    const labels = labelsOf(d);
    const box = h('span', { class: 'row' });
    if (d.n?.trim()) box.append(h('span', { class: 'noteicon', title: 'Has a note' }, '📝'));
    for (const l of labels.slice(0, MAX_PILLS)) box.append(pill(l));
    if (labels.length > MAX_PILLS) box.append(h('span', { class: 'more' }, `+${labels.length - MAX_PILLS}`));
    setContent(host.shadowRoot!, box);
  }
}
