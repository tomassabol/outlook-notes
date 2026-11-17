import baseCss from '../shared/base.css?inline';
import pageCss from './page.css?inline';
import css from './editor.css?inline';
import { h, injectStyle, pill } from '../shared/dom';
import { convKey, friendlyError, getConv, getLabels, newId, saveConv, saveLabels } from '../shared/store';
import { LABEL_COLORS, QUICK_EMOJI, type ConvData, type Label } from '../shared/types';

const params = new URLSearchParams(location.search);
const convId = params.get('conv') ?? '';
const subject = params.get('subject') ?? '';
const focus = params.get('focus') === 'labels' ? 'labels' : 'note';
document.documentElement.dataset.theme = params.get('theme') === 'dark' ? 'dark' : 'light';
injectStyle(document, baseCss + pageCss + css);

const isMac = navigator.platform.toLowerCase().includes('mac');
const close = () => parent.postMessage({ type: 'onl-close' }, '*');

async function main(): Promise<void> {
  if (!convId) return close();
  const key = convKey(convId);
  let [labels, data] = await Promise.all([getLabels(), getConv(key)]);
  const selected = new Set((data?.l ?? []).filter((id) => labels.some((l) => l.id === id)));

  const textarea = h('textarea', { id: 'note', rows: 6, placeholder: 'Write a private note…', spellcheck: true });
  textarea.value = data?.n ?? '';
  const chips = h('div', { class: 'chips', role: 'group', 'aria-label': 'Labels' });
  const error = h('p', { class: 'error', role: 'alert', hidden: true });
  const showError = (e: unknown) => {
    error.textContent = friendlyError(e);
    error.hidden = false;
  };

  function renderChips(): void {
    chips.replaceChildren();
    if (!labels.length) chips.append(h('span', { class: 'muted' }, 'No labels yet. Create one →'));
    for (const l of labels) {
      const on = selected.has(l.id);
      const chip = h(
        'button',
        {
          class: 'chip',
          type: 'button',
          'aria-pressed': String(on),
          onclick: () => {
            if (selected.has(l.id)) selected.delete(l.id);
            else selected.add(l.id);
            renderChips();
            (chips.querySelector(`[data-id="${l.id}"]`) as HTMLElement | null)?.focus();
          },
        },
        h('span', { class: 'check', 'aria-hidden': 'true' }, on ? '✓' : ''),
        pill(l),
      );
      chip.dataset.id = l.id;
      chips.append(chip);
    }
  }

  // ----- new label form -----
  let color = LABEL_COLORS[(labels.length + 5) % LABEL_COLORS.length];
  const emojiInput = h('input', { class: 'input emoji', maxlength: 8, 'aria-label': 'Emoji', placeholder: '🙂' });
  const textInput = h('input', { class: 'input text', maxlength: 40, 'aria-label': 'Label text', placeholder: 'Label text' });
  const swatches = h('div', { class: 'swatches', role: 'radiogroup', 'aria-label': 'Color' });
  const previewSlot = h('span');
  const updatePreview = () => {
    const l: Label = { id: '', emoji: emojiInput.value.trim(), text: textInput.value.trim() || 'Preview', color };
    previewSlot.replaceChildren(pill(l));
  };
  const renderSwatches = () => {
    swatches.replaceChildren(
      ...LABEL_COLORS.map((c) => {
        const b = h('button', {
          class: 'swatch',
          type: 'button',
          role: 'radio',
          'aria-checked': String(c === color),
          'aria-label': c,
          onclick: () => {
            color = c;
            renderSwatches();
            updatePreview();
          },
        });
        b.style.background = c;
        return b;
      }),
    );
  };
  emojiInput.addEventListener('input', updatePreview);
  textInput.addEventListener('input', updatePreview);

  const newForm = h(
    'div',
    { class: 'newlabel', hidden: true },
    h('div', { class: 'line' }, emojiInput, textInput),
    h(
      'div',
      { class: 'quick', 'aria-label': 'Quick emoji' },
      ...QUICK_EMOJI.map((e) =>
        h('button', { type: 'button', title: e, onclick: () => { emojiInput.value = e; updatePreview(); textInput.focus(); } }, e),
      ),
    ),
    swatches,
    h(
      'div',
      { class: 'preview' },
      previewSlot,
      h('span', { class: 'spacer', style: 'flex:1' }),
      h('button', { class: 'btn small', type: 'button', onclick: () => toggleNew(false) }, 'Cancel'),
      h('button', { class: 'btn small primary', type: 'button', onclick: () => void addLabel() }, 'Add label'),
    ),
  );

  function toggleNew(open: boolean): void {
    newForm.hidden = !open;
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
    color = LABEL_COLORS[(labels.length + 5) % LABEL_COLORS.length];
    renderChips();
    toggleNew(false);
  }

  const newBtn = h('button', { class: 'btn link', type: 'button', onclick: () => toggleNew(newForm.hidden) }, '+ New label');

  // ----- save / clear -----
  async function save(): Promise<void> {
    const n = textarea.value.replace(/\s+$/, '');
    const l = labels.filter((x) => selected.has(x.id)).map((x) => x.id);
    const next: ConvData = { s: (subject || data?.s || '').slice(0, 60), u: Date.now() };
    if (n.trim()) next.n = n;
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
    { class: 'dialog', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'title' },
    h('h1', { id: 'title' }, h('span', { class: 'h-ic', 'aria-hidden': 'true' }, '📝'), 'Note & labels'),
    h('p', { class: 'subject', title: subject }, subject || 'Conversation'),
    h('label', { class: 'field-label', for: 'note' }, 'Note'),
    textarea,
    h('div', { class: 'section-head' }, h('span', { class: 'field-label' }, 'Labels'), newBtn),
    chips,
    newForm,
    error,
    h(
      'footer',
      {},
      data ? h('button', { class: 'btn danger', type: 'button', onclick: () => void clearAll() }, 'Clear all') : null,
      h('span', { class: 'spacer' }),
      h('span', { class: 'hint muted' }, `${isMac ? '⌘' : 'Ctrl'}+Enter to save`),
      h('button', { class: 'btn', type: 'button', onclick: close }, 'Cancel'),
      h('button', { class: 'btn primary', type: 'button', onclick: () => void save() }, 'Save'),
    ),
  );
  const backdrop = h('div', { class: 'backdrop' }, dialog);
  backdrop.addEventListener('mousedown', (e) => {
    if (e.target === backdrop) close();
  });
  document.body.append(backdrop);
  renderChips();

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      if (!newForm.hidden) toggleNew(false);
      else close();
    } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      void save();
    } else if (e.key === 'Enter' && !newForm.hidden && newForm.contains(e.target as Node) && e.target instanceof HTMLInputElement) {
      e.preventDefault();
      void addLabel();
    }
  });

  if (focus === 'labels') (chips.querySelector('.chip') as HTMLElement | null ?? newBtn).focus();
  else {
    textarea.focus();
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
  }
}

void main();
