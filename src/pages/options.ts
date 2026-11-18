import tokensCss from '../shared/tokens.css?inline';
import fallbackCss from './fluent-fallback.css?inline';
import pageCss from './page.css?inline';
import css from './options.css?inline';
import { h, injectStyle, pill } from '../shared/dom';
import { icon } from '../shared/icons';
import { PALETTE_NAMES, applyPaletteVars, paletteName, paletteTitle } from '../shared/palette';
import {
  CONV_PREFIX, LABELS_KEY, deleteLabel, friendlyError, getAllConvs, getLabels, isConvData, isLabel, migrateLabels, newId, saveConv, saveLabels, usage,
} from '../shared/store';
import type { ConvData, ExportFile, Label } from '../shared/types';

const dark = matchMedia('(prefers-color-scheme: dark)');
const applyTheme = () => (document.documentElement.dataset.theme = dark.matches ? 'dark' : 'light');
applyTheme();
dark.addEventListener('change', applyTheme);
injectStyle(document, tokensCss + fallbackCss + pageCss + css);

let labels: Label[] = [];
let convs: Record<string, ConvData> = {};
let query = '';
let openPalette: string | null = null;

// ---------- layout ----------
const labelsList = h('div');
const labelsCount = h('span', { class: 'badge' });
const notesList = h('div', { class: 'notes' });
const notesCount = h('span', { class: 'badge' });
const progress = h('div', { class: 'progress', role: 'progressbar', 'aria-label': 'Sync storage used', 'aria-valuemin': '0', 'aria-valuemax': '100' }, h('div'));
const usageText = h('p', { class: 'caption', style: 'margin:0' });
const status = h('p', { class: 'status', role: 'status' });
const fileInput = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });

const say = (text: string, isError = false) => {
  status.textContent = text;
  status.className = isError ? 'status error-text' : 'status';
};

const btn = (cls: string, label: string, ic: Parameters<typeof icon>[0] | null, onClick: () => void, extra: Record<string, string | boolean> = {}) =>
  h('button', { class: `btn ${cls}`.trim(), type: 'button', onclick: onClick as EventListener, ...extra }, ic ? icon(ic) : null, label);

document.body.append(
  h(
    'main',
    {},
    h(
      'header',
      { class: 'app-header' },
      h('span', { class: 'app-icon' }, icon('note24')),
      h('div', {}, h('h1', {}, 'Outlook Notes & Labels'), h('p', { class: 'caption' }, 'Private notes and labels for Outlook on the web, synced with your Chrome profile.')),
    ),
    h(
      'section',
      { class: 'card', 'aria-labelledby': 'h-labels' },
      h('div', { class: 'card-head' }, h('h2', { id: 'h-labels' }, 'Labels'), labelsCount, h('span', { class: 'spacer' }), btn('small subtle', 'New label', 'add16', () => void addLabel())),
      h('p', { class: 'caption' }, 'Changes apply everywhere the label is used.'),
      labelsList,
    ),
    h(
      'section',
      { class: 'card', 'aria-labelledby': 'h-notes' },
      h('div', { class: 'card-head' }, h('h2', { id: 'h-notes' }, 'Notes & labeled emails'), notesCount),
      h('p', { class: 'caption' }, 'Newest first.'),
      h(
        'div',
        { class: 'searchbox' },
        icon('search20'),
        h('input', {
          class: 'input',
          type: 'search',
          placeholder: 'Search subjects, notes and labels',
          'aria-label': 'Search notes',
          oninput: ((e: Event) => {
            query = (e.target as HTMLInputElement).value.toLowerCase();
            renderNotes();
          }) as EventListener,
        }),
      ),
      notesList,
    ),
    h(
      'section',
      { class: 'card', 'aria-labelledby': 'h-storage' },
      h('div', { class: 'card-head' }, h('h2', { id: 'h-storage' }, 'Storage & backup')),
      progress,
      usageText,
      h(
        'div',
        { class: 'toolbar' },
        btn('', 'Export', 'download20', () => void exportData()),
        btn('', 'Import…', 'upload20', () => fileInput.click()),
        h('span', { class: 'spacer' }),
        btn('subtle danger', 'Delete everything', 'delete20', () => void wipe()),
      ),
      fileInput,
      status,
    ),
  ),
);

// ---------- labels ----------
let saveTimer: number | undefined;
function saveLabelsSoon(): void {
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => saveLabels(labels).catch((e) => say(friendlyError(e), true)), 400);
}

function renderLabels(): void {
  labelsCount.textContent = String(labels.length);
  labelsList.replaceChildren();
  if (!labels.length) labelsList.append(h('p', { class: 'empty-state' }, 'No labels yet.'));
  labels.forEach((l, i) => {
    const preview = h('span', { class: 'preview' });
    const refresh = () => preview.replaceChildren(pill(l));
    refresh();
    const emoji = h('input', { class: 'input emoji', value: l.emoji, maxlength: 8, 'aria-label': 'Emoji' });
    const text = h('input', { class: 'input', value: l.text, maxlength: 40, 'aria-label': 'Label name' });
    emoji.addEventListener('input', () => { l.emoji = emoji.value.trim(); refresh(); saveLabelsSoon(); });
    text.addEventListener('input', () => { l.text = text.value.trim(); refresh(); saveLabelsSoon(); });

    const colorBtn = h('button', {
      class: 'swatch color-btn',
      type: 'button',
      title: `Color: ${paletteTitle(paletteName(l.color))}`,
      'aria-label': `Color: ${paletteTitle(paletteName(l.color))}`,
      'aria-expanded': String(openPalette === l.id),
      onclick: () => {
        openPalette = openPalette === l.id ? null : l.id;
        renderLabels();
      },
    });
    applyPaletteVars(colorBtn, l.color);

    const move = (d: number) => () => {
      [labels[i], labels[i + d]] = [labels[i + d], labels[i]];
      renderLabels();
      saveLabelsSoon();
    };
    const used = Object.values(convs).filter((c) => c.l?.includes(l.id)).length;
    const row = h(
      'div',
      { class: 'label-row' },
      emoji,
      text,
      colorBtn,
      preview,
      h(
        'div',
        { class: 'actions' },
        h('button', { class: 'btn subtle icon-only', type: 'button', title: 'Move up', 'aria-label': 'Move up', disabled: i === 0, onclick: move(-1) }, icon('arrowUp16')),
        h('button', { class: 'btn subtle icon-only', type: 'button', title: 'Move down', 'aria-label': 'Move down', disabled: i === labels.length - 1, onclick: move(1) }, icon('arrowDown16')),
        h(
          'button',
          {
            class: 'btn subtle icon-only danger',
            type: 'button',
            title: 'Delete label',
            'aria-label': `Delete label ${l.text}`,
            onclick: async () => {
              const msg = used
                ? `Delete “${l.emoji} ${l.text}”? It’s used on ${used} email${used === 1 ? '' : 's'} and will be removed from them.`
                : `Delete “${l.emoji} ${l.text}”?`;
              if (!confirm(msg)) return;
              clearTimeout(saveTimer);
              await saveLabels(labels).catch(() => undefined);
              await deleteLabel(l.id).catch((e) => say(friendlyError(e), true));
            },
          },
          icon('delete16'),
        ),
      ),
    );
    if (openPalette === l.id) {
      const current = paletteName(l.color);
      row.append(
        h(
          'div',
          { class: 'palette swatches', role: 'radiogroup', 'aria-label': 'Label color' },
          ...PALETTE_NAMES.map((c) => {
            const b = h(
              'button',
              {
                class: 'swatch',
                type: 'button',
                role: 'radio',
                'aria-checked': String(c === current),
                title: paletteTitle(c),
                'aria-label': paletteTitle(c),
                onclick: () => {
                  l.color = c;
                  openPalette = null;
                  renderLabels();
                  saveLabelsSoon();
                },
              },
              icon('checkmark12'),
            );
            applyPaletteVars(b, c);
            return b;
          }),
        ),
      );
    }
    labelsList.append(row);
  });
}

async function addLabel(): Promise<void> {
  labels.push({ id: newId(), emoji: '🏷️', text: 'New label', color: PALETTE_NAMES[(labels.length * 7 + 3) % PALETTE_NAMES.length] });
  renderLabels();
  try {
    await saveLabels(labels);
  } catch (e) {
    say(friendlyError(e), true);
  }
  const inputs = labelsList.querySelectorAll<HTMLInputElement>('.label-row input:not(.emoji)');
  inputs[inputs.length - 1]?.select();
}

// ---------- notes ----------
function renderNotes(): void {
  const entries = Object.entries(convs).sort((a, b) => b[1].u - a[1].u);
  notesCount.textContent = String(entries.length);
  const byId = new Map(labels.map((l) => [l.id, l]));
  const matches = entries.filter(([, d]) => {
    if (!query) return true;
    const ls = (d.l ?? []).map((id) => byId.get(id)).map((l) => (l ? `${l.emoji} ${l.text}` : '')).join(' ');
    return `${d.s} ${d.n ?? ''} ${ls}`.toLowerCase().includes(query);
  });
  notesList.replaceChildren(
    ...matches.map(([key, d]) => {
      const ls = (d.l ?? []).map((id) => byId.get(id)).filter((l): l is Label => !!l);
      return h(
        'article',
        { class: 'note-item' },
        h('span', { class: 'lead' }, icon(d.n ? 'note20' : 'tag20')),
        h(
          'div',
          { class: 'main' },
          h(
            'div',
            { class: 'top' },
            h('span', { class: 'subj', title: d.s }, d.s || '(no subject)'),
            h('span', { class: 'caption' }, new Date(d.u).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' })),
          ),
          ls.length ? h('div', { class: 'pills' }, ...ls.map((l) => pill(l))) : null,
          d.n ? h('div', { class: 'text' }, d.n) : null,
        ),
        h(
          'button',
          {
            class: 'btn subtle icon-only danger',
            type: 'button',
            title: 'Delete note and labels',
            'aria-label': `Delete note for ${d.s}`,
            onclick: async () => {
              if (confirm(`Delete the note and labels for “${d.s || 'this email'}”?`)) await saveConv(key, null).catch((e) => say(friendlyError(e), true));
            },
          },
          icon('delete16'),
        ),
      );
    }),
  );
  if (!entries.length) notesList.append(h('p', { class: 'empty-state' }, 'No notes yet. In Outlook, open an email and choose “Add note”, or right-click an email in the list.'));
  else if (!matches.length) notesList.append(h('p', { class: 'empty-state' }, 'Nothing matches your search.'));
}

async function renderUsage(): Promise<void> {
  const { used, quota, items } = await usage();
  const pct = Math.min(100, (used / quota) * 100);
  (progress.firstElementChild as HTMLElement).style.width = `${Math.max(pct, 0.5)}%`;
  progress.className = `progress${pct > 90 ? ' full' : pct > 70 ? ' warn' : ''}`;
  progress.setAttribute('aria-valuenow', String(Math.round(pct)));
  usageText.textContent = `${(used / 1024).toFixed(1)} KB of ${(quota / 1024).toFixed(0)} KB Chrome sync storage used · ${items} item${items === 1 ? '' : 's'}`;
}

// ---------- backup ----------
async function exportData(): Promise<void> {
  const file: ExportFile = {
    app: 'outlook-notes',
    version: 1,
    exportedAt: new Date().toISOString(),
    labels: await getLabels(),
    conversations: await getAllConvs(),
  };
  const url = URL.createObjectURL(new Blob([JSON.stringify(file, null, 2)], { type: 'application/json' }));
  const a = h('a', { href: url, download: `outlook-notes-${new Date().toISOString().slice(0, 10)}.json` });
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  say('Exported.');
}

fileInput.addEventListener('change', async () => {
  const f = fileInput.files?.[0];
  fileInput.value = '';
  if (!f) return;
  try {
    const raw = JSON.parse(await f.text()) as Partial<ExportFile>;
    if (raw.app !== 'outlook-notes' || raw.version !== 1) throw new Error('This isn’t an Outlook Notes export file.');
    const inLabels = (raw.labels ?? []).filter(isLabel).map((l) => ({ ...l, color: paletteName(l.color) }));
    const inConvs = Object.entries(raw.conversations ?? {}).filter(([k, v]) => k.startsWith(CONV_PREFIX) && isConvData(v));
    if (!confirm(`Import ${inLabels.length} labels and ${inConvs.length} notes? Existing entries with the same ID are overwritten; others are kept.`)) return;
    const current = await getLabels();
    const merged = [...current.filter((l) => !inLabels.some((x) => x.id === l.id)), ...inLabels];
    await chrome.storage.sync.set({ [LABELS_KEY]: merged, ...Object.fromEntries(inConvs) });
    say(`Imported ${inLabels.length} labels and ${inConvs.length} notes.`);
  } catch (e) {
    say(`Import failed: ${friendlyError(e)}`, true);
  }
});

async function wipe(): Promise<void> {
  if (!confirm('Delete ALL notes and labels? This can’t be undone. Export a backup first if you might need them.')) return;
  await chrome.storage.sync.clear();
  say('All data deleted.');
}

// ---------- load + live updates ----------
async function load(): Promise<void> {
  await migrateLabels().catch(() => false);
  [labels, convs] = await Promise.all([getLabels(), getAllConvs()]);
  renderLabels();
  renderNotes();
  await renderUsage();
}

chrome.storage.onChanged.addListener(async (changes, area) => {
  if (area !== 'sync') return;
  const editingLabels = labelsList.contains(document.activeElement) && document.activeElement instanceof HTMLInputElement;
  convs = await getAllConvs();
  if (LABELS_KEY in changes && !editingLabels) {
    labels = await getLabels();
    renderLabels();
  }
  renderNotes();
  await renderUsage();
});

void load();
