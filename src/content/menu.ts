import { h } from '../shared/dom';
import { icon, type IconName } from '../shared/icons';
import type { EditorFocus, Msg } from '../shared/types';
import { openEditor } from './editor-frame';
import { MENU_SELECTOR, rowFromTarget } from './selectors';
import { dataFor } from './state';
import { currentOpenConvId, subjectFor } from './target';
import { makeHost, setContent, toast } from './ui';

let lastContext: { convId: string; at: number } | null = null;

/** Close Outlook's (Fluent UI) context menu after one of our items is chosen. */
function closeOutlookMenu(menu: HTMLElement): void {
  const opts = { key: 'Escape', code: 'Escape', bubbles: true, cancelable: true };
  const target = (document.activeElement as HTMLElement | null) ?? menu;
  target.dispatchEvent(new KeyboardEvent('keydown', opts));
  if (menu.isConnected) {
    // Fallback: an outside click dismisses Fluent callouts.
    for (const type of ['pointerdown', 'mousedown', 'mouseup', 'click']) {
      document.body.dispatchEvent(new MouseEvent(type, { bubbles: true, clientX: 1, clientY: 1 }));
    }
  }
}

function inject(menu: HTMLElement, convId: string): void {
  // Outlook's menu is a div[role=menu] of div items; older Fluent menus use <ul><li>.
  // <li> can't host a shadow root, so in that case wrap a <div> host in a plain list item.
  const list = menu.matches('ul') ? menu : menu.querySelector(':scope > ul') ?? menu;
  const { host: inner, root } = makeHost('data-onl-menu-inner', 'div');
  let host: HTMLElement = inner;
  if (list.tagName === 'UL') {
    host = document.createElement('li');
    host.setAttribute('role', 'presentation');
    host.setAttribute('data-onl', '');
    host.style.listStyle = 'none';
    host.append(inner);
  }
  host.setAttribute('data-onl-menu', '');
  const d = dataFor(convId);
  const item = (ic: IconName, text: string, focus: EditorFocus) =>
    h(
      'button',
      {
        class: 'mi',
        type: 'button',
        role: 'menuitem',
        onclick: ((e: Event) => {
          e.preventDefault();
          e.stopPropagation();
          closeOutlookMenu(menu);
          openEditor(convId, subjectFor(convId), focus);
        }) as EventListener,
      },
      h('span', { class: 'ic' }, icon(ic)),
      h('span', { class: 'label' }, text),
    );
  setContent(
    root,
    h(
      'div',
      { class: 'menu' },
      item(d?.n?.trim() ? 'note20' : 'noteAdd20', d?.n?.trim() ? 'Edit note…' : 'Add note…', 'note'),
      item('tag20', 'Labels…', 'labels'),
      h('div', { class: 'divider', role: 'separator' }),
    ),
  );
  for (const ev of ['mousedown', 'pointerdown', 'click'] as const) host.addEventListener(ev, (e) => e.stopPropagation());
  list.prepend(host);
}

/** Outlook renders its menu asynchronously after the contextmenu event; wait briefly for it. */
function injectWhenMenuAppears(convId: string): void {
  const tryInject = (): boolean => {
    const menus = [...document.querySelectorAll<HTMLElement>(MENU_SELECTOR)];
    const menu = menus.at(-1);
    if (!menu) return false;
    if (!menu.querySelector('[data-onl-menu]')) inject(menu, convId);
    return true;
  };
  const obs = new MutationObserver(() => {
    if (tryInject()) obs.disconnect();
  });
  obs.observe(document.body, { childList: true, subtree: true });
  window.setTimeout(() => obs.disconnect(), 1500);
}

export function initMenu(): void {
  document.addEventListener(
    'contextmenu',
    (e) => {
      const hit = rowFromTarget(e.target);
      lastContext = hit ? { convId: hit.convId, at: Date.now() } : null;
      if (hit) injectWhenMenuAppears(hit.convId);
    },
    true,
  );

  // Chrome's own context menu (shown where Outlook doesn't override right-click).
  chrome.runtime.onMessage.addListener((msg: Msg) => {
    if (msg?.type !== 'open-editor') return;
    const recent = lastContext && Date.now() - lastContext.at < 60_000 ? lastContext.convId : null;
    const convId = recent ?? currentOpenConvId();
    if (!convId) {
      toast('Open an email or right-click one in the list first.');
      return;
    }
    openEditor(convId, subjectFor(convId), msg.focus);
  });
}
