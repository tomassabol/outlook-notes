import { hasContent } from '../shared/store';
import { FIELD_NAMES, fieldBox, h, pill } from '../shared/dom';
import { icon, type IconName } from '../shared/icons';
import type { EditorFocus } from '../shared/types';
import { openEditor } from './editor-frame';
import { findBannerAnchor, findReadingPane } from './selectors';
import { dataFor, labelsOf, signature, state } from './state';
import { currentOpenConvId, subjectFor } from './target';
import { copyable, labelCopyText, makeHost, setContent } from './ui';

const ATTR = 'data-onl-banner';
const COLLAPSE_KEY = 'bannerCollapsed';
let collapsed = false;

export async function initBanner(): Promise<void> {
  const r = await chrome.storage.local.get(COLLAPSE_KEY);
  collapsed = !!r[COLLAPSE_KEY];
}

function button(text: string, ic: IconName, title: string, onClick: () => void, iconOnly = false): HTMLButtonElement {
  return h(
    'button',
    { class: iconOnly ? 'btn icon-only' : 'btn', type: 'button', title, 'aria-label': title, onclick: onClick as EventListener },
    icon(ic),
    iconOnly ? null : text,
  );
}

/** Shows the note bar for the open conversation; removes it when nothing is open. */
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
    // Clicks inside the bar shouldn't reach Outlook's reading-pane handlers.
    for (const ev of ['click', 'mousedown', 'dblclick'] as const) host.addEventListener(ev, (e) => e.stopPropagation());
    anchor.after(host);
  }
  host.dataset.sig = sig;
  host.dataset.theme = state.theme;

  const edit = (focus: EditorFocus) => () => openEditor(convId, subjectFor(convId), focus);

  if (!hasContent(d)) {
    setContent(
      host.shadowRoot!,
      h(
        'div',
        { class: 'empty' },
        button('Add note', 'noteAdd16', 'Add a private note', edit('note')),
        button('Add label', 'tag16', 'Add labels', edit('labels')),
      ),
    );
    return;
  }

  const labels = labelsOf(d);
  const note = d.n?.trim() ?? '';
  const isCollapsed = collapsed && !!note;
  const toggle = () => {
    collapsed = !collapsed;
    void chrome.storage.local.set({ [COLLAPSE_KEY]: collapsed });
    updateBanner();
  };

  const labelPills = labels.length ? h('span', { class: 'pills' }, ...labels.map((l) => copyable(pill(l), labelCopyText(l)))) : null;
  const fieldEls = (['b', 'w'] as const)
    .map((k) => [k, d[k]?.trim()] as const)
    .filter(([, v]) => !!v)
    .map(([k, v]) => copyable(fieldBox(v!, FIELD_NAMES[k]), v!));
  const fieldsRow = fieldEls.length ? h('span', { class: 'fields' }, ...fieldEls) : null;
  // Head: "Note" when there's a note; otherwise the first of labels / fields.
  const headContent = note ? h('span', { class: 'title' }, 'Note') : (labelPills ?? fieldsRow);
  const rest = [labelPills, fieldsRow].filter((x): x is HTMLSpanElement => !!x && x !== headContent);
  const head = h(
    'div',
    { class: 'head' },
    headContent,
    isCollapsed ? h('span', { class: 'peek' }, note.split('\n')[0]) : null,
  );
  const actions = h(
    'div',
    { class: 'actions' },
    button(note ? 'Edit' : 'Add note', note ? 'edit16' : 'noteAdd16', note ? 'Edit note' : 'Add note', edit('note')),
    button('Labels', 'tag16', 'Edit labels', edit('labels')),
    note ? button('', isCollapsed ? 'chevronDown16' : 'chevronUp16', isCollapsed ? 'Expand note' : 'Collapse note', toggle, true) : null,
  );

  setContent(
    host.shadowRoot!,
    h(
      'div',
      { class: 'bar', role: 'note', 'aria-label': 'Your note' },
      h('div', { class: 'top' }, h('span', { class: 'lead' }, icon(note ? 'note20' : 'tag20')), head, actions),
      rest.length ? h('div', { class: 'indent' }, ...rest) : null,
      note && !isCollapsed ? h('div', { class: 'text' }, note) : null,
    ),
  );
}
