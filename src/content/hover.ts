import { hasContent } from '../shared/store';
import { h, pill } from '../shared/dom';
import { icon, type IconName } from '../shared/icons';
import type { EditorFocus } from '../shared/types';
import { openEditor } from './editor-frame';
import { rowFromTarget } from './selectors';
import { subjectFor } from './target';
import { dataFor, labelsOf } from './state';
import { makeHost, setContent } from './ui';

const SHOW_DELAY = 400;
/** Time to move the pointer from the row onto the card (or back) before it closes. */
const HIDE_GRACE = 300;

let currentRow: HTMLElement | null = null;
let showTimer: number | undefined;
let hideTimer: number | undefined;
let pointerX = 0;
let card: HTMLElement | null = null;
let overCard = false;
/** Mouse button went down inside the card (a text selection may be in progress). */
let selecting = false;

/** True while the user has text selected inside the card, so it stays open for copying. */
function hasSelection(): boolean {
  if (!card?.shadowRoot) return false;
  const root = card.shadowRoot as ShadowRoot & { getSelection?: () => Selection | null };
  const sel = root.getSelection?.() ?? document.getSelection();
  return !!sel && !sel.isCollapsed && sel.toString().trim().length > 0 && !!sel.anchorNode && card.shadowRoot.contains(sel.anchorNode);
}

function hide(): void {
  clearTimeout(showTimer);
  clearTimeout(hideTimer);
  card?.remove();
  card = null;
  overCard = false;
  selecting = false;
}

function scheduleHide(): void {
  clearTimeout(hideTimer);
  hideTimer = window.setTimeout(() => {
    if (overCard || selecting || hasSelection()) return;
    currentRow = null;
    hide();
  }, HIDE_GRACE);
}

function show(row: HTMLElement, convId: string): void {
  const d = dataFor(convId);
  if (!hasContent(d) || !row.isConnected) return;
  hide();
  const { host, root } = makeHost('data-onl-card', 'div');
  const labels = labelsOf(d);
  const note = d.n?.trim();
  setContent(
    root,
    h(
      'div',
      { class: 'card', role: 'tooltip' },
      h(
        'div',
        { class: 'caption' },
        icon(note ? 'note16' : 'tag16'),
        h('span', { class: 'spacer' }, note ? 'Note' : 'Labels'),
        button(note ? 'Edit' : 'Add note', note ? 'edit16' : 'noteAdd16', note ? 'Edit note' : 'Add note', convId, 'note'),
        button('', 'tag16', 'Edit labels', convId, 'labels', true),
      ),
      labels.length ? clickToEdit(h('div', { class: 'pills editable', title: 'Click to edit labels' }, ...labels.map((l) => pill(l))), convId, 'labels') : null,
      note ? clickToEdit(h('div', { class: 'note editable', title: 'Click to edit' }, note), convId, 'note') : null,
    ),
  );

  host.addEventListener('mouseenter', () => {
    overCard = true;
    clearTimeout(hideTimer);
  });
  host.addEventListener('mouseleave', (e) => {
    overCard = false;
    // Moving back onto the row keeps it open; anywhere else closes after the grace period.
    if (!(currentRow && currentRow.contains(e.relatedTarget as Node | null))) scheduleHide();
  });
  host.addEventListener('mousedown', (e) => {
    selecting = true;
    e.stopPropagation(); // don't let Outlook treat this as a click elsewhere
  });
  for (const ev of ['click', 'dblclick', 'contextmenu'] as const) host.addEventListener(ev, (e) => e.stopPropagation());

  host.style.visibility = 'hidden';
  document.body.append(host);
  card = host;

  const r = row.getBoundingClientRect();
  const { width, height } = host.getBoundingClientRect();
  const x = Math.min(Math.max(8, pointerX - 24), innerWidth - width - 8);
  let y = r.bottom + 2;
  if (y + height > innerHeight - 8) y = Math.max(8, r.top - height - 2);
  host.style.left = `${x}px`;
  host.style.top = `${y}px`;
  host.style.visibility = '';
}

function edit(convId: string, focus: EditorFocus): void {
  currentRow = null;
  hide();
  openEditor(convId, subjectFor(convId), focus);
}

/**
 * A plain click opens the editor; a drag (text selection for copying) doesn't. A click that
 * merely collapses an existing selection counts as a click.
 */
function clickToEdit<T extends HTMLElement>(el: T, convId: string, focus: EditorFocus): T {
  let down: { x: number; y: number; hadSelection: boolean } | null = null;
  el.addEventListener('mousedown', (e) => {
    down = e.button === 0 ? { x: e.clientX, y: e.clientY, hadSelection: hasSelection() } : null;
  });
  el.addEventListener('click', (e) => {
    const start = down;
    down = null;
    if (!start || e.detail > 1) return; // double/triple click selects words/lines
    const moved = Math.hypot(e.clientX - start.x, e.clientY - start.y) > 4;
    if (moved || (hasSelection() && !start.hadSelection)) return;
    edit(convId, focus);
  });
  return el;
}

function button(text: string, ic: IconName, title: string, convId: string, focus: EditorFocus, iconOnly = false): HTMLButtonElement {
  return h(
    'button',
    {
      class: iconOnly ? 'btn icon-only' : 'btn',
      type: 'button',
      title,
      'aria-label': title,
      onclick: (() => edit(convId, focus)) as EventListener,
    },
    icon(ic),
    iconOnly ? null : text,
  );
}

const inCard = (e: Event) => !!card && e.composedPath().includes(card);

export function initHover(): void {
  document.addEventListener('mousemove', (e) => (pointerX = e.clientX), { passive: true, capture: true });

  document.addEventListener(
    'mouseover',
    (e) => {
      if (inCard(e)) return;
      const hit = rowFromTarget(e.target);
      if (!hit) return;
      if (hit.row === currentRow) {
        clearTimeout(hideTimer); // came back to the same row
        return;
      }
      // Keep a card with selected text until the user clicks elsewhere.
      if (card && (selecting || hasSelection())) return;
      currentRow = hit.row;
      hide();
      if (!hasContent(dataFor(hit.convId))) return;
      showTimer = window.setTimeout(() => show(hit.row, hit.convId), SHOW_DELAY);
    },
    true,
  );

  document.addEventListener(
    'mouseout',
    (e) => {
      if (!currentRow || inCard(e)) return;
      const to = e.relatedTarget as Node | null;
      if (currentRow.contains(to) || (card && to && card.contains(to))) return;
      clearTimeout(showTimer);
      if (card) scheduleHide();
      else currentRow = null;
    },
    true,
  );

  document.addEventListener(
    'mouseup',
    () => {
      if (!selecting) return;
      selecting = false;
      if (!overCard && !hasSelection()) scheduleHide();
    },
    true,
  );

  // Clicking or scrolling elsewhere closes it; Escape closes it; other keys (e.g. ⌘C) don't.
  document.addEventListener('mousedown', (e) => { if (card && !inCard(e)) { currentRow = null; hide(); } }, { capture: true });
  document.addEventListener('wheel', (e) => { if (card && !inCard(e)) { currentRow = null; hide(); } }, { passive: true, capture: true });
  document.addEventListener('contextmenu', (e) => { if (card && !inCard(e)) { currentRow = null; hide(); } }, { capture: true });
  document.addEventListener('keydown', (e) => { if (card && e.key === 'Escape') { currentRow = null; hide(); } }, { capture: true });
}
