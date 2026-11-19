import { hasContent } from '../shared/store';
import { FIELD_NAMES, fieldBox, h, pill } from '../shared/dom';
import type { EditorFocus } from '../shared/types';
import { openEditor } from './editor-frame';
import { rowFromTarget } from './selectors';
import { subjectFor } from './target';
import { dataFor, labelsOf } from './state';
import { copyable, labelCopyText, makeHost, setContent } from './ui';

const SHOW_DELAY = 400;
/** Time to move the pointer from the row onto the card (or back) before it closes. */
const HIDE_GRACE = 700;
/**
 * While a card is open, the pointer must *rest* this long on another row before the card
 * switches to it. Moving across rows on the way to the card (the usual diagonal path) doesn't
 * count, because every mousemove restarts the timer.
 */
const SWITCH_DELAY = 450;
/** The card overlaps the row by this much so there's no dead gap between them. */
const OVERLAP = 4;

let currentRow: HTMLElement | null = null;
let showTimer: number | undefined;
let hideTimer: number | undefined;
let switchTimer: number | undefined;
/** Another row the pointer is over while a card is open (a possible switch target). */
let pending: { row: HTMLElement; convId: string } | null = null;
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
  clearTimeout(switchTimer);
  pending = null;
  card?.remove();
  card = null;
  overCard = false;
  selecting = false;
}

function scheduleHide(): void {
  clearTimeout(hideTimer);
  hideTimer = window.setTimeout(() => {
    // Over another row: the switch timer decides instead.
    if (overCard || selecting || hasSelection() || pending) return;
    currentRow = null;
    hide();
  }, HIDE_GRACE);
}

/** (Re)start the timer that switches the open card to the row the pointer is resting on. */
function restartSwitch(): void {
  clearTimeout(switchTimer);
  switchTimer = window.setTimeout(() => {
    const target = pending;
    if (!target || overCard || selecting || hasSelection()) return;
    pending = null;
    currentRow = target.row;
    hide();
    if (hasContent(dataFor(target.convId))) show(target.row, target.convId);
  }, SWITCH_DELAY);
}

function show(row: HTMLElement, convId: string): void {
  const d = dataFor(convId);
  if (!hasContent(d) || !row.isConnected) return;
  hide();
  const { host, root } = makeHost('data-onl-card', 'div');
  const labels = labelsOf(d);
  const note = d.n?.trim();
  const fields = (['b', 'w'] as const)
    .map((k) => [k, d[k]?.trim()] as const)
    .filter(([, v]) => !!v)
    .map(([k, v]) => copyable(fieldBox(v!, FIELD_NAMES[k]), v!));
  const hasTop = labels.length > 0 || fields.length > 0;
  // Clicking the card (anywhere except a label or field, which copy) opens the editor.
  const body = clickToEdit(
    h(
      'div',
      { class: 'card', role: 'tooltip' },
      labels.length ? h('div', { class: 'pills' }, ...labels.map((l) => copyable(pill(l), labelCopyText(l)))) : null,
      fields.length ? h('div', { class: 'fields stack' }, ...fields) : null,
      note && hasTop ? h('div', { class: 'divider', role: 'separator' }) : null,
      note ? h('div', { class: 'note editable', title: 'Click to edit' }, note) : null,
    ),
    convId,
    'note',
  );
  setContent(root, body);

  host.addEventListener('mouseenter', () => {
    overCard = true;
    clearTimeout(hideTimer);
    clearTimeout(switchTimer);
    pending = null;
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
  let y = r.bottom - OVERLAP;
  if (y + height > innerHeight - 8) y = Math.max(8, r.top - height + OVERLAP);
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

const inCard = (e: Event) => !!card && e.composedPath().includes(card);

export function initHover(): void {
  document.addEventListener(
    'mousemove',
    (e) => {
      pointerX = e.clientX;
      // Still moving over another row: not a deliberate stop yet, so keep postponing the switch.
      if (pending && pending.row.contains(e.target as Node)) restartSwitch();
    },
    { passive: true, capture: true },
  );

  document.addEventListener(
    'mouseover',
    (e) => {
      if (inCard(e)) return;
      const hit = rowFromTarget(e.target);
      if (!hit) {
        // Off the rows (and not on the card): forget any switch target; the grace period runs.
        if (pending) {
          pending = null;
          clearTimeout(switchTimer);
          if (card && !overCard) scheduleHide();
        }
        return;
      }
      if (hit.row === currentRow) {
        // Back on the card's own row.
        clearTimeout(hideTimer);
        clearTimeout(switchTimer);
        pending = null;
        return;
      }
      if (card) {
        // A card is open: the pointer may just be passing over this row on its way to the card.
        if (selecting || hasSelection()) return; // keep a card with selected text
        pending = hit;
        restartSwitch();
        return;
      }
      currentRow = hit.row;
      clearTimeout(showTimer);
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
