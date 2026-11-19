import tokensCss from '../shared/tokens.css?inline';
import fallbackCss from './fluent-fallback.css?inline';
import pageCss from './page.css?inline';
import css from './editor.css?inline';
import { FIELD_NAMES, h, injectStyle, pill } from '../shared/dom';
import { icon } from '../shared/icons';
import { PALETTE_NAMES, applyPaletteVars, paletteTitle } from '../shared/palette';
import { convKey, friendlyError, getConv, getLabels, newId, saveConv, saveLabels } from '../shared/store';
import { QUICK_EMOJI, type ConvData, type Label } from '../shared/types';

const params = new URLSearchParams(location.search);
const convId = params.get('conv') ?? '';
const subject = params.get('subject') ?? '';
const focus = params.get('focus') === 'labels' ? 'labels' : 'note';
const root = document.documentElement;
root.dataset.theme = params.get('theme') === 'dark' ? 'dark' : 'light';
injectStyle(document, tokensCss + fallbackCss + pageCss + css);

const isMac = navigator.platform.toLowerCase().includes('mac');
const close = () => parent.postMessage({ type: 'onl-close' }, '*');

// Adopt Outlook's live theme tokens (sent by the content script) so the dialog matches exactly.
let markRevealed: () => void;
const revealed = new Promise<void>((res) => (markRevealed = res));
const reveal = () => {
  root.classList.add('ready');
  markRevealed();
};
window.addEventListener('message', (e) => {
  if (e.source !== parent) return;
  const data = e.data as { type?: string; tokens?: Record<string, string> };
  if (data?.type !== 'onl-tokens' || !data.tokens) return;
  for (const [k, v] of Object.entries(data.tokens)) if (k.startsWith('--')) root.style.setProperty(k, v);
  root.dataset.tokens = '';
  reveal();
});
parent.postMessage({ type: 'onl-ready' }, '*');
setTimeout(reveal, 250);

async function main(): Promise<void> {
  if (!convId) return close();
  const key = convKey(convId);
  let [labels, data] = await Promise.all([getLabels(), getConv(key)]);
  const selected = new Set((data?.l ?? []).filter((id) => labels.some((l) => l.id === id)));

  const textarea = h('textarea', { id: 'note', class: 'input note', rows: 6, placeholder: 'Write a private note…', spellcheck: true });
  textarea.value = data?.n ?? '';
  const bNumberInput = h('input', { id: 'bnumber', class: 'input', maxlength: 40, placeholder: 'B12345', autocomplete: 'off', spellcheck: false });
  const workOrderInput = h('input', { id: 'workorder', class: 'input', maxlength: 40, placeholder: 'WO-12345', autocomplete: 'off', spellcheck: false });
  bNumberInput.value = data?.b ?? '';
  workOrderInput.value = data?.w ?? '';
  const tags = h('div', { class: 'tags', role: 'group', 'aria-labelledby': 'labels-label' });
  const error = h('p', { class: 'error', role: 'alert', hidden: true });
  const showError = (e: unknown) => {
    error.replaceChildren(icon('info16'), h('span', {}, friendlyError(e)));
    error.hidden = false;
  };

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

  // ----- new label form -----
  const nextColor = () => PALETTE_NAMES[(labels.length * 7 + 3) % PALETTE_NAMES.length];
  let color = nextColor();
  const emojiInput = h('input', { class: 'input emoji', maxlength: 8, 'aria-label': 'Emoji', placeholder: '🙂' });
  const textInput = h('input', { class: 'input text', maxlength: 40, 'aria-label': 'Label name', placeholder: 'Label name' });
  const swatches = h('div', { class: 'swatches', role: 'radiogroup', 'aria-label': 'Color' });
  const previewSlot = h('span');
  const updatePreview = () => {
    previewSlot.replaceChildren(pill({ id: '', emoji: emojiInput.value.trim(), text: textInput.value.trim() || 'Preview', color }));
  };
  const renderSwatches = () => {
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
  };
  emojiInput.addEventListener('input', updatePreview);
  textInput.addEventListener('input', updatePreview);

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
    try {
      labels = [...(await getLabels()), label];
      await saveLabels(labels);
    } catch (e) {
      return showError(e);
    }
    selected.add(label.id);
    emojiInput.value = '';
    textInput.value = '';
    color = nextColor();
    renderTags();
    toggleNew(false);
  }

  const newBtn = h(
    'button',
    { class: 'btn small subtle', type: 'button', 'aria-expanded': 'false', onclick: () => toggleNew(newForm.hidden) },
    icon('add16'),
    'New label',
  );

  // ----- save / clear -----
  async function save(): Promise<void> {
    const n = textarea.value.replace(/\s+$/, '');
    const l = labels.filter((x) => selected.has(x.id)).map((x) => x.id);
    const next: ConvData = { s: (subject || data?.s || '').slice(0, 60), u: Date.now() };
    if (n.trim()) next.n = n;
    const b = bNumberInput.value.trim();
    const w = workOrderInput.value.trim();
    if (b) next.b = b;
    if (w) next.w = w;
    if (l.length) next.l = l;
    try {
      await saveConv(key, next);
      close();
    } catch (e) {
      showError(e);
    }
  }

  async function clearAll(): Promise<void> {
    if (!confirm('Remove the note and all labels from this conversation?')) return;
    try {
      await saveConv(key, null);
      close();
    } catch (e) {
      showError(e);
    }
  }

  const dialog = h(
    'div',
    { class: 'dialog', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'title', 'aria-describedby': 'subject', tabindex: '-1' },
    h(
      'header',
      {},
      h('div', { class: 'titles' }, h('h1', { id: 'title' }, data ? 'Edit note & labels' : 'Add note & labels'), h('p', { class: 'subject', id: 'subject', title: subject }, subject || 'Conversation')),
      h('button', { class: 'btn subtle icon-only close', type: 'button', title: 'Close', 'aria-label': 'Close', onclick: close }, icon('dismiss20')),
    ),
    h(
      'div',
      { class: 'body' },
      h(
        'div',
        { class: 'ids' },
        h('div', {}, h('label', { class: 'field-label', for: 'bnumber' }, FIELD_NAMES.b), bNumberInput),
        h('div', {}, h('label', { class: 'field-label', for: 'workorder' }, FIELD_NAMES.w), workOrderInput),
      ),
      h('div', {}, h('label', { class: 'field-label', for: 'note' }, 'Note'), textarea),
      h(
        'div',
        {},
        h('div', { class: 'section-head' }, h('span', { class: 'field-label', id: 'labels-label' }, 'Labels'), newBtn),
        tags,
        newForm,
      ),
      error,
    ),
    h(
      'footer',
      {},
      data ? h('button', { class: 'btn subtle danger', type: 'button', onclick: () => void clearAll() }, icon('delete20'), 'Remove all') : null,
      h('span', { class: 'spacer' }),
      h('button', { class: 'btn', type: 'button', onclick: close }, 'Cancel'),
      h('button', { class: 'btn primary', type: 'button', title: `Save (${isMac ? '⌘' : 'Ctrl'}+Enter)`, onclick: () => void save() }, 'Save'),
    ),
  );
  const backdrop = h('div', { class: 'backdrop' }, dialog);
  backdrop.addEventListener('mousedown', (e) => {
    if (e.target === backdrop) close();
  });
  document.body.append(backdrop);
  renderTags();

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      if (!newForm.hidden) toggleNew(false);
      else close();
    } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      void save();
    } else if (e.key === 'Enter' && !newForm.hidden && e.target instanceof HTMLInputElement && newForm.contains(e.target)) {
      e.preventDefault();
      void addLabel();
    }
  });

  // Hidden elements can't take focus, so wait until the themed dialog is shown.
  await revealed;
  window.focus();
  // For labels, focus the dialog rather than the first tag: a programmatic focus would draw a
  // focus ring on that tag. Tab still moves through the tags from here.
  if (focus === 'labels') dialog.focus();
  else {
    textarea.focus();
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
  }
}

void main();
