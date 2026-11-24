/**
 * Outlook add-in task pane: note, B number, work order and labels for the selected email's
 * conversation. Data is kept in this computer's add-in storage (localStorage). Labels are also
 * mirrored to Outlook categories on the email, so they show in the message list.
 */
import tokensCss from '../shared/tokens.css?inline';
import fallbackCss from '../pages/fluent-fallback.css?inline';
import pageCss from '../pages/page.css?inline';
import css from './taskpane.css?inline';
import { FIELD_NAMES, h, injectStyle, pill } from '../shared/dom';
import { icon } from '../shared/icons';
import { PALETTE_NAMES, applyPaletteVars, paletteName, paletteTitle } from '../shared/palette';
import {
  CONV_PREFIX, DEFAULT_LABELS, LABELS_KEY, convKey, friendlyError, getAllConvs, getConv, getLabels, isConvData, isLabel,
  newId, saveConv, saveLabels, setStorageArea,
} from '../shared/store';
import { QUICK_EMOJI, type ConvData, type ExportFile, type Label } from '../shared/types';
import { categoryColor, categoryName, mockHost, officeHost, type CurrentItem, type Host } from './host';
import { localArea } from './local-area';

setStorageArea(localArea);
injectStyle(document, tokensCss + fallbackCss + pageCss + css);

const root = document.documentElement;
const isMock = new URLSearchParams(location.search).has('mock');

// ---------- theme ----------
function isDarkColor(hex: string | undefined): boolean | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!m) return null;
  const n = parseInt(m[1], 16);
  const lum = 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255);
  return lum < 128;
}
const darkQuery = matchMedia('(prefers-color-scheme: dark)');
function applyTheme(): void {
  let dark: boolean | null = null;
  try {
    dark = isDarkColor((Office.context as { officeTheme?: Office.OfficeTheme } | undefined)?.officeTheme?.bodyBackgroundColor);
  } catch {
    // Office not loaded (mock)
  }
  root.dataset.theme = (dark ?? darkQuery.matches) ? 'dark' : 'light';
}
darkQuery.addEventListener('change', applyTheme);

// ---------- state ----------
let host: Host;
let cur: CurrentItem | null = null;
let key = '';
let labels: Label[] = [];
let data: ConvData | undefined;
const selected = new Set<string>();
let saveTimer: ReturnType<typeof setTimeout> | undefined;
let loadToken = 0;
/** Category names known to exist in each mailbox's master list. */
const knownMaster = new Map<string, Set<string>>();

// ---------- elements ----------
const subjectEl = h('p', { class: 'subject' });
const textarea = h('textarea', { id: 'note', class: 'input note', rows: 6, placeholder: 'Write a private note…', spellcheck: true });
const bInput = h('input', { id: 'bnumber', class: 'input', maxlength: 40, placeholder: 'B12345', autocomplete: 'off', spellcheck: false });
const wInput = h('input', { id: 'workorder', class: 'input', maxlength: 40, placeholder: 'WO-12345', autocomplete: 'off', spellcheck: false });
const tags = h('div', { class: 'tags', role: 'group', 'aria-labelledby': 'labels-label' });
const status = h('span', { class: 'caption status', role: 'status' });
const warn = h('p', { class: 'warn', role: 'alert', hidden: true });
const error = h('p', { class: 'error', role: 'alert', hidden: true });
const removeBtn = h('button', { class: 'btn small subtle danger', type: 'button', onclick: () => void clearAll() }, icon('delete16'), 'Remove all');
const fileInput = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });

function showError(e: unknown): void {
  error.replaceChildren(icon('info16'), h('span', {}, friendlyError(e)));
  error.hidden = false;
}
function showWarn(msg: string | null): void {
  warn.hidden = !msg;
  if (msg) warn.replaceChildren(icon('info16'), h('span', {}, msg));
}
const say = (t: string) => (status.textContent = t);

// ---------- saving ----------
function collect(): ConvData {
  const next: ConvData = { s: (cur?.subject || data?.s || '').slice(0, 60), u: Date.now() };
  const n = textarea.value.replace(/\s+$/, '');
  const b = bInput.value.trim();
  const w = wInput.value.trim();
  const l = labels.filter((x) => selected.has(x.id)).map((x) => x.id);
  if (n.trim()) next.n = n;
  if (b) next.b = b;
  if (w) next.w = w;
  if (l.length) next.l = l;
  return next;
}

async function saveNow(): Promise<void> {
  clearTimeout(saveTimer);
  saveTimer = undefined;
  if (!key) return;
  const next = collect();
  try {
    await saveConv(key, next);
    data = next;
    error.hidden = true;
    say('Saved on this computer');
    removeBtn.hidden = !(next.n || next.b || next.w || next.l);
  } catch (e) {
    showError(e);
  }
}

function saveSoon(): void {
  say('Saving…');
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => void saveNow(), 600);
}
for (const el of [textarea, bInput, wInput]) el.addEventListener('input', saveSoon);
window.addEventListener('pagehide', () => void saveNow());

// ---------- Outlook categories ----------
async function ensureMaster(names: { name: string; color: string }[]): Promise<void> {
  if (!cur) return;
  let known = knownMaster.get(cur.mailbox);
  if (!known || names.some((n) => !known!.has(n.name.toLowerCase()))) {
    known = new Set((await host.masterCategories()).map((n) => n.toLowerCase()));
    knownMaster.set(cur.mailbox, known);
  }
  const missing = names.filter((n) => !known!.has(n.name.toLowerCase()));
  if (missing.length) {
    await host.addMasterCategories(missing);
    for (const m of missing) known.add(m.name.toLowerCase());
  }
}

/**
 * Makes the open email's categories match the conversation's labels. Only categories named
 * after one of our labels are touched; any other categories on the email are left alone.
 */
async function syncCategories(): Promise<void> {
  if (!cur) return;
  const token = loadToken;
  try {
    const want = labels.filter((l) => selected.has(l.id));
    const wantNames = new Set(want.map(categoryName));
    const managed = new Set(labels.map(categoryName));
    const have = await host.itemCategories();
    if (token !== loadToken) return;
    const toAdd = want.filter((l) => !have.includes(categoryName(l)));
    const toRemove = have.filter((n) => managed.has(n) && !wantNames.has(n));
    if (toAdd.length) {
      await ensureMaster(toAdd.map((l) => ({ name: categoryName(l), color: categoryColor(l) })));
      if (token !== loadToken) return;
      await host.addItemCategories(toAdd.map(categoryName));
    }
    if (toRemove.length && token === loadToken) await host.removeItemCategories(toRemove);
    showWarn(null);
  } catch (e) {
    if (token === loadToken) showWarn(`Couldn’t update Outlook categories (${friendlyError(e)}). Labels are still saved here.`);
  }
}

// ---------- labels ----------
function renderTags(): void {
  tags.replaceChildren();
  if (!labels.length) tags.append(h('span', { class: 'empty-tags' }, 'No labels yet.'));
  for (const l of labels) {
    const on = selected.has(l.id);
    const tag = h(
      'button',
      {
        class: 'tag',
        type: 'button',
        'aria-pressed': String(on),
        title: on ? `Remove “${l.text}”` : `Add “${l.text}”`,
        onclick: () => {
          if (selected.has(l.id)) selected.delete(l.id);
          else selected.add(l.id);
          renderTags();
          (tags.querySelector(`[data-id="${l.id}"]`) as HTMLElement | null)?.focus();
          void saveNow().then(syncCategories);
        },
      },
      h('span', { class: 'check' }, icon('checkmark12')),
      l.emoji ? h('span', { 'aria-hidden': 'true' }, l.emoji) : null,
      l.text,
    );
    applyPaletteVars(tag, l.color);
    tag.dataset.id = l.id;
    tags.append(tag);
  }
}

const nextColor = () => PALETTE_NAMES[(labels.length * 7 + 3) % PALETTE_NAMES.length];
let color = nextColor();
const emojiInput = h('input', { class: 'input emoji', maxlength: 8, 'aria-label': 'Emoji', placeholder: '🙂' });
const textInput = h('input', { class: 'input text', maxlength: 40, 'aria-label': 'Label name', placeholder: 'Label name' });
const swatches = h('div', { class: 'swatches', role: 'radiogroup', 'aria-label': 'Color' });
const previewSlot = h('span');
const updatePreview = () =>
  previewSlot.replaceChildren(pill({ id: '', emoji: emojiInput.value.trim(), text: textInput.value.trim() || 'Preview', color }));
function renderSwatches(): void {
  swatches.replaceChildren(
    ...PALETTE_NAMES.map((c) => {
      const b = h(
        'button',
        {
          class: 'swatch',
          type: 'button',
          role: 'radio',
          'aria-checked': String(c === color),
          'aria-label': paletteTitle(c),
          title: paletteTitle(c),
          onclick: () => {
            color = c;
            renderSwatches();
            updatePreview();
            (swatches.querySelector('[aria-checked="true"]') as HTMLElement | null)?.focus();
          },
        },
        icon('checkmark12'),
      );
      applyPaletteVars(b, c);
      return b;
    }),
  );
}
emojiInput.addEventListener('input', updatePreview);
textInput.addEventListener('input', updatePreview);

const newBtn = h(
  'button',
  { class: 'btn small subtle', type: 'button', 'aria-expanded': 'false', onclick: () => toggleNew(newForm.hidden) },
  icon('add16'),
  'New label',
);
const newForm = h(
  'div',
  { class: 'newlabel', hidden: true, role: 'group', 'aria-label': 'New label' },
  h('div', { class: 'line' }, emojiInput, textInput),
  h(
    'div',
    { class: 'quick', 'aria-label': 'Suggested emoji' },
    ...QUICK_EMOJI.map((e) =>
      h('button', { type: 'button', title: e, onclick: () => { emojiInput.value = e; updatePreview(); textInput.focus(); } }, e),
    ),
  ),
  swatches,
  h(
    'div',
    { class: 'preview-row' },
    previewSlot,
    h('span', { class: 'spacer' }),
    h('button', { class: 'btn small', type: 'button', onclick: () => toggleNew(false) }, 'Cancel'),
    h('button', { class: 'btn small primary', type: 'button', onclick: () => void addLabel() }, 'Create'),
  ),
);
newForm.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target instanceof HTMLInputElement) {
    e.preventDefault();
    void addLabel();
  } else if (e.key === 'Escape') toggleNew(false);
});

function toggleNew(open: boolean): void {
  newForm.hidden = !open;
  newBtn.setAttribute('aria-expanded', String(open));
  if (open) {
    renderSwatches();
    updatePreview();
    emojiInput.focus();
  } else newBtn.focus();
}

async function addLabel(): Promise<void> {
  const emoji = emojiInput.value.trim();
  const text = textInput.value.trim();
  if (!emoji && !text) return textInput.focus();
  const label: Label = { id: newId(), emoji, text, color };
  if (labels.some((l) => categoryName(l).toLowerCase() === categoryName(label).toLowerCase())) {
    return showError('A label with that name already exists.');
  }
  try {
    labels = [...(await getLabels()), label];
    await saveLabels(labels);
  } catch (e) {
    return showError(e);
  }
  emojiInput.value = '';
  textInput.value = '';
  color = nextColor();
  toggleNew(false);
  if (cur) {
    selected.add(label.id);
    await saveNow();
    void syncCategories();
  }
  renderTags();
}

async function clearAll(): Promise<void> {
  if (!confirm('Remove the note and all labels from this conversation?')) return;
  textarea.value = bInput.value = wInput.value = '';
  selected.clear();
  renderTags();
  await saveNow();
  void syncCategories();
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
  const json = JSON.stringify(file, null, 2);
  const url = URL.createObjectURL(new Blob([json], { type: 'application/json' }));
  h('a', { href: url, download: `outlook-notes-${new Date().toISOString().slice(0, 10)}.json` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  // Some Outlook web views ignore downloads; the clipboard copy is the fallback.
  await navigator.clipboard?.writeText(json).catch(() => undefined);
  say('Exported (also copied to the clipboard)');
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
    await localArea.set({ [LABELS_KEY]: [...current.filter((l) => !inLabels.some((x) => x.id === l.id)), ...inLabels], ...Object.fromEntries(inConvs) });
    say(`Imported ${inLabels.length} labels and ${inConvs.length} notes`);
    await load();
  } catch (e) {
    showError(`Import failed: ${friendlyError(e)}`);
  }
});

// ---------- layout ----------
const empty = h('div', { class: 'empty-state', hidden: true }, icon('note24'), h('p', {}, 'Select an email to see or add its note.'));
const form = h(
  'main',
  { class: 'pane' },
  subjectEl,
  h(
    'div',
    { class: 'ids' },
    h('div', {}, h('label', { class: 'field-label', for: 'bnumber' }, FIELD_NAMES.b), bInput),
    h('div', {}, h('label', { class: 'field-label', for: 'workorder' }, FIELD_NAMES.w), wInput),
  ),
  h('div', {}, h('label', { class: 'field-label', for: 'note' }, 'Note'), textarea),
  h('div', {}, h('div', { class: 'section-head' }, h('span', { class: 'field-label', id: 'labels-label' }, 'Labels'), newBtn), tags, newForm),
  warn,
  error,
  h('div', { class: 'row' }, status, h('span', { class: 'spacer' }), removeBtn),
);
const footer = h(
  'footer',
  { class: 'pane-footer' },
  h('span', { class: 'caption' }, 'Backup'),
  h('button', { class: 'btn small subtle', type: 'button', onclick: () => void exportData() }, icon('download20'), 'Export'),
  h('button', { class: 'btn small subtle', type: 'button', onclick: () => fileInput.click() }, icon('upload20'), 'Import'),
  fileInput,
);

// ---------- loading the selected email ----------
async function load(): Promise<void> {
  if (saveTimer) await saveNow(); // finish the previous email's pending edit first
  const token = ++loadToken;
  cur = host.current();
  error.hidden = true;
  showWarn(null);
  toggleNew(false);
  if (!cur) {
    key = '';
    form.hidden = true;
    empty.hidden = false;
    return;
  }
  const item = cur;
  const k = convKey(item.convId);
  const [ls, d] = await Promise.all([getLabels(), getConv(k)]);
  if (token !== loadToken) return;
  key = k;
  labels = ls;
  data = d;
  selected.clear();
  for (const id of d?.l ?? []) if (labels.some((l) => l.id === id)) selected.add(id);
  subjectEl.textContent = item.subject || 'Conversation';
  subjectEl.title = item.subject;
  textarea.value = d?.n ?? '';
  bInput.value = d?.b ?? '';
  wInput.value = d?.w ?? '';
  removeBtn.hidden = !d;
  say(d ? `Last edited ${new Date(d.u).toLocaleString()}` : '');
  form.hidden = false;
  empty.hidden = true;
  renderTags();
  void syncCategories();
}

async function start(outlook: Host): Promise<void> {
  host = outlook;
  applyTheme();
  // First run: seed the default labels (same as the Chrome extension).
  if (!(LABELS_KEY in (await localArea.get(LABELS_KEY)))) await saveLabels(DEFAULT_LABELS);
  document.body.append(form, empty, footer);
  host.onItemChanged(() => void load());
  await load();
}

if (isMock) {
  void start(mockHost());
} else {
  Office.onReady(() => void start(officeHost()));
}
