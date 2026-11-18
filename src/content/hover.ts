import { hasContent } from '../shared/store';
import { h, pill } from '../shared/dom';
import { icon } from '../shared/icons';
import { rowFromTarget } from './selectors';
import { dataFor, labelsOf } from './state';
import { makeHost, setContent } from './ui';

const DELAY = 400;
let currentRow: HTMLElement | null = null;
let timer: number | undefined;
let pointerX = 0;
let card: HTMLElement | null = null;

function hide(): void {
  clearTimeout(timer);
  card?.remove();
  card = null;
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
      h('div', { class: 'caption' }, icon(note ? 'note16' : 'tag16'), note ? 'Note' : 'Labels'),
      labels.length ? h('div', { class: 'pills' }, ...labels.map((l) => pill(l))) : null,
      note ? h('div', { class: `note${note.length > 360 || note.split('\n').length > 9 ? ' long' : ''}` }, note) : null,
    ),
  );
  host.style.visibility = 'hidden';
  document.body.append(host);
  card = host;

  const r = row.getBoundingClientRect();
  const { width, height } = host.getBoundingClientRect();
  const x = Math.min(Math.max(8, pointerX - 24), innerWidth - width - 8);
  let y = r.bottom + 4;
  if (y + height > innerHeight - 8) y = Math.max(8, r.top - height - 4);
  host.style.left = `${x}px`;
  host.style.top = `${y}px`;
  host.style.visibility = '';
}

export function initHover(): void {
  document.addEventListener('mousemove', (e) => (pointerX = e.clientX), { passive: true, capture: true });

  document.addEventListener(
    'mouseover',
    (e) => {
      const hit = rowFromTarget(e.target);
      if (!hit || hit.row === currentRow) return;
      currentRow = hit.row;
      hide();
      if (!hasContent(dataFor(hit.convId))) return;
      timer = window.setTimeout(() => show(hit.row, hit.convId), DELAY);
    },
    true,
  );

  document.addEventListener(
    'mouseout',
    (e) => {
      if (currentRow && !currentRow.contains(e.relatedTarget as Node | null)) {
        currentRow = null;
        hide();
      }
    },
    true,
  );

  for (const ev of ['mousedown', 'wheel', 'keydown', 'contextmenu'] as const) {
    document.addEventListener(ev, hide, { passive: true, capture: true });
  }
}

