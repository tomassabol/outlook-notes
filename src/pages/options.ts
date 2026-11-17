import baseCss from '../shared/base.css?inline';
import pageCss from './page.css?inline';
import css from './options.css?inline';
import { h, injectStyle, pill } from '../shared/dom';
import {
  CONV_PREFIX, LABELS_KEY, deleteLabel, friendlyError, getAllConvs, getLabels, isConvData, isLabel, newId, saveConv, saveLabels, usage,
} from '../shared/store';
import { LABEL_COLORS, type ConvData, type ExportFile, type Label } from '../shared/types';

const dark = matchMedia('(prefers-color-scheme: dark)');
const applyTheme = () => (document.documentElement.dataset.theme = dark.matches ? 'dark' : 'light');
applyTheme();
dark.addEventListener('change', applyTheme);
injectStyle(document, baseCss + pageCss + css);

let labels: Label[] = [];
let convs: Record<string, ConvData> = {};
let query = '';

// ---------- layout ----------
const labelsList = h('div');
const labelsCount = h('span', { class: 'count' });
const notesList = h('div', { class: 'notes' });
const notesCount = h('span', { class: 'count' });
const meter = h('div', { class: 'meter', role: 'meter', 'aria-label': 'Sync storage used' }, h('div'));
const usageText = h('p', { class: 'muted', style: 'margin:6px 0 0' });
const status = h('p', { class: 'status', role: 'status' });
const fileInput = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });

const say = (text: string, isError = false) => {
  status.textContent = text;
  status.className = isError ? 'status error' : 'status muted';
};

document.body.append(
  h(
    'main',
    {},
    h('header', {}, h('h1', {}, '📝 Outlook Notes & Labels'), h('p', { class: 'muted' }, 'Private notes and labels for Outlook on the web, synced through your Chrome profile.')),
    h(
      'section',
      { 'aria-labelledby': 'h-labels' },
      h('h2', { id: 'h-labels' }, 'Labels', labelsCount),
      labelsList,
      h('div', { class: 'toolbar' }, h('button', { class: 'btn', type: 'button', onclick: () => void addLabel() }, '+ Add label')),
    ),
    h(
      'section',
      { 'aria-labelledby': 'h-notes' },
      h('h2', { id: 'h-notes' }, 'Notes & labeled emails', notesCount),
      h('input', {
        class: 'input search',
        type: 'search',
        placeholder: 'Search subject, note or label…',
        'aria-label': 'Search notes',
        oninput: ((e: Event) => {
          query = (e.target as HTMLInputElement).value.toLowerCase();
          renderNotes();
        }) as EventListener,
      }),
      notesList,
    ),
    h(
      'section',
      { 'aria-labelledby': 'h-storage' },
      h('h2', { id: 'h-storage' }, 'Storage & backup'),
      meter,
      usageText,
      h(
        'div',
        { class: 'toolbar' },
        h('button', { class: 'btn', type: 'button', onclick: () => void exportData() }, 'Export JSON'),
        h('button', { class: 'btn', type: 'button', onclick: () => fileInput.click() }, 'Import JSON…'),
        h('span', { style: 'flex:1' }),
        h('button', { class: 'btn danger', type: 'button', onclick: () => void wipe() }, 'Delete everything'),
      ),
      fileInput,
      status,
    ),
  ),
);

// ---------- labels ----------
const pendingSaves = new Map<string, number>();
function saveLabelsSoon(): void {
  clearTimeout(pendingSaves.get('labels'));
  pendingSaves.set(
    'labels',
    window.setTimeout(() => saveLabels(labels).catch((e) => say(friendlyError(e), true)), 400),
  );
}

function renderLabels(): void {
  labelsCount.textContent = String(labels.length);
  labelsList.replaceChildren();
  if (!labels.length) labelsList.append(h('p', { class: 'muted' }, 'No labels yet.'));
  labels.forEach((l, i) => {
    const preview = h('span');
    const refresh = () => preview.replaceChildren(pill(l));
    refresh();
    const emoji = h('input', { class: 'input emoji', value: l.emoji, maxlength: 8, 'aria-label': 'Emoji' });
    const text = h('input', { class: 'input', value: l.text, maxlength: 40, 'aria-label': 'Label text' });
    const color = h('input', { type: 'color', value: l.color, 'aria-label': 'Color' });
    emoji.addEventListener('input', () => { l.emoji = emoji.value.trim(); refresh(); saveLabelsSoon(); });
    text.addEventListener('input', () => { l.text = text.value.trim(); refresh(); saveLabelsSoon(); });
    color.addEventListener('input', () => { l.color = color.value; refresh(); saveLabelsSoon(); });
    const move = (d: number) => () => {
      [labels[i], labels[i + d]] = [labels[i + d], labels[i]];
      renderLabels();
      saveLabelsSoon();
    };
    const used = Object.values(convs).filter((c) => c.l?.includes(l.id)).length;
    labelsList.append(
      h(
        'div',
        { class: 'label-row' },
        emoji,
        text,
        color,
        h(
          'div',
          { class: 'actions' },
          preview,
          h('button', { class: 'icon-btn', type: 'button', title: 'Move up', 'aria-label': 'Move up', disabled: i === 0, onclick: move(-1) }, '↑'),
          h('button', { class: 'icon-btn', type: 'button', title: 'Move down', 'aria-label': 'Move down', disabled: i === labels.length - 1, onclick: move(1) }, '↓'),
          h(
            'button',
            {
              class: 'icon-btn danger',
              type: 'button',
              title: 'Delete label',
              'aria-label': `Delete label ${l.text}`,
              onclick: async () => {
                const msg = used ? `Delete “${l.emoji} ${l.text}”? It’s used on ${used} email${used === 1 ? '' : 's'} and will be removed from them.` : `Delete “${l.emoji} ${l.text}”?`;
                if (!confirm(msg)) return;
                clearTimeout(pendingSaves.get('labels'));
                await saveLabels(labels).catch(() => undefined);
                await deleteLabel(l.id).catch((e) => say(friendlyError(e), true));
              },
            },
            '✕',
          ),
        ),
      ),
    );
  });
}

async function addLabel(): Promise<void> {
  labels.push({ id: newId(), emoji: '🏷️', text: 'New label', color: LABEL_COLORS[labels.length % LABEL_COLORS.length] });
  renderLabels();
  try {
    await saveLabels(labels);
  } catch (e) {
    say(friendlyError(e), true);
  }
  const inputs = labelsList.querySelectorAll<HTMLInputElement>('.label-row input:not(.emoji):not([type=color])');
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
        { class: 'note-card' },
        h(
          'div',
          { class: 'top' },
          h('span', { class: 'subj', title: d.s }, d.s || '(no subject)'),
          h('span', { class: 'date' }, new Date(d.u).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' })),
          h(
            'button',
            {
              class: 'icon-btn danger',
              type: 'button',
              title: 'Delete note and labels',
              'aria-label': `Delete note for ${d.s}`,
              onclick: async () => {
                if (confirm(`Delete the note and labels for “${d.s || 'this email'}”?`)) await saveConv(key, null).catch((e) => say(friendlyError(e), true));
              },
            },
            '✕',
          ),
        ),
        ls.length ? h('div', { class: 'pills' }, ...ls.map(pill)) : null,
        d.n ? h('div', { class: 'text' }, d.n) : null,
      );
    }),
  );
  if (!entries.length) notesList.append(h('p', { class: 'muted' }, 'No notes yet. In Outlook, open an email and click “📝 Add note”, or right-click an email in the list.'));
  else if (!matches.length) notesList.append(h('p', { class: 'muted' }, 'Nothing matches your search.'));
}

async function renderUsage(): Promise<void> {
  const { used, quota, items } = await usage();
  const pct = Math.min(100, (used / quota) * 100);
  (meter.firstElementChild as HTMLElement).style.width = `${pct}%`;
  meter.className = `meter${pct > 90 ? ' full' : pct > 70 ? ' warn' : ''}`;
  meter.setAttribute('aria-valuenow', String(Math.round(pct)));
  usageText.textContent = `${(used / 1024).toFixed(1)} KB of ${(quota / 1024).toFixed(0)} KB Chrome sync storage used (${Math.round(pct)}%), ${items} item${items === 1 ? '' : 's'}.`;
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
    const inLabels = (raw.labels ?? []).filter(isLabel);
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
  [labels, convs] = await Promise.all([getLabels(), getAllConvs()]);
  renderLabels();
  renderNotes();
  await renderUsage();
}

chrome.storage.onChanged.addListener(async (changes, area) => {
  if (area !== 'sync') return;
  const editingLabels = labelsList.contains(document.activeElement);
  convs = await getAllConvs();
  if (LABELS_KEY in changes && !editingLabels) {
    labels = await getLabels();
    renderLabels();
  }
  renderNotes();
  await renderUsage();
});

void load();
