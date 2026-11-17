import { hasContent } from '../shared/store';
import { h, pill } from '../shared/dom';
import type { EditorFocus } from '../shared/types';
import { openEditor } from './editor-frame';
import { findBannerAnchor, findReadingPane } from './selectors';
import { dataFor, labelsOf, signature, state } from './state';
import { currentOpenConvId, subjectFor } from './target';
import { makeHost, setContent } from './ui';

const ATTR = 'data-onl-banner';
const COLLAPSE_KEY = 'bannerCollapsed';
let collapsed = false;

export async function initBanner(): Promise<void> {
  const r = await chrome.storage.local.get(COLLAPSE_KEY);
  collapsed = !!r[COLLAPSE_KEY];
}

/** Shows the sticky-note banner for the open conversation; removes it when nothing is open. */
export function updateBanner(): void {
  const convId = currentOpenConvId();
  const pane = convId ? findReadingPane() : null;
  // Outlook keeps the last row selected after closing an email, so also require a subject heading.
  const anchor = pane ? findBannerAnchor(pane) : null;
  let host = document.querySelector<HTMLElement>(`[${ATTR}]`);
  if (!convId || !pane || !anchor || !pane.contains(anchor)) {
    host?.remove();
    return;
  }
  const d = dataFor(convId);
  const sig = signature(convId, d, String(collapsed));
  if (host && host.previousElementSibling === anchor && host.dataset.sig === sig) return;
  if (!host || host.previousElementSibling !== anchor) {
    host?.remove();
    host = makeHost(ATTR, 'div').host;
    // Clicks inside the banner shouldn't reach Outlook's reading-pane handlers.
    for (const ev of ['click', 'mousedown', 'dblclick'] as const) host.addEventListener(ev, (e) => e.stopPropagation());
    anchor.after(host);
  }
  host.dataset.sig = sig;
  host.dataset.theme = state.theme;

  const edit = (focus: EditorFocus) => () => openEditor(convId, subjectFor(convId), focus);
  const btn = (text: string, title: string, onClick: () => void, cls = 'btn', icon = '') =>
    h('button', { class: cls, type: 'button', title, onclick: onClick as EventListener }, icon ? h('span', { class: 'ic' }, icon) : null, text);

  if (!hasContent(d)) {
    setContent(
      host.shadowRoot!,
      h('div', { class: 'empty' }, btn('Add note', 'Add a private note', edit('note'), 'btn', '📝'), btn('Add label', 'Add labels', edit('labels'), 'btn', '🏷')),
    );
    return;
  }

  const labels = labelsOf(d);
  const note = d.n?.trim() ?? '';
  const toggle = () => {
    collapsed = !collapsed;
    void chrome.storage.local.set({ [COLLAPSE_KEY]: collapsed });
    updateBanner();
  };
  const head = h(
    'div',
    { class: 'head' },
    note ? h('span', { class: 'title' }, '📝') : null,
    labels.length ? h('span', { class: 'pills' }, ...labels.map(pill)) : null,
    collapsed && note ? h('span', { class: 'peek' }, note.split('\n')[0]) : h('span', { class: 'spacer' }),
    btn(note ? 'Edit note' : 'Add note', 'Edit note', edit('note')),
    btn('Labels', 'Edit labels', edit('labels'), 'btn', '🏷'),
    note ? btn(collapsed ? '▸' : '▾', collapsed ? 'Expand note' : 'Collapse note', toggle, 'btn icon') : null,
  );
  setContent(
    host.shadowRoot!,
    h('div', { class: 'banner', role: 'note', 'aria-label': 'Your note' }, head, !collapsed && note ? h('div', { class: 'body' }, note) : null),
  );
}
