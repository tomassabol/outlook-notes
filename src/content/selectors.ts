/**
 * Every assumption about Outlook's DOM lives here. Outlook on the web uses generated class
 * names, so we only rely on ARIA roles and data attributes. If Microsoft changes the DOM,
 * this is the file to fix.
 */

export interface RowHit {
  row: HTMLElement;
  convId: string;
}

const CONV_ATTR = 'data-convid';
const OUR = '[data-onl]';

function rowOf(el: Element): HTMLElement {
  return (el.closest('[role="option"]') as HTMLElement | null) ?? (el as HTMLElement);
}

/** All conversation rows currently rendered in the message list. */
export function findRows(): RowHit[] {
  const out: RowHit[] = [];
  const seen = new Set<Element>();
  for (const el of document.querySelectorAll(`[${CONV_ATTR}]`)) {
    const convId = el.getAttribute(CONV_ATTR);
    if (!convId) continue;
    const row = rowOf(el);
    if (seen.has(row)) continue;
    seen.add(row);
    out.push({ row, convId });
  }
  return out;
}

/** The conversation row that contains an event target (if any). */
export function rowFromTarget(t: EventTarget | null): RowHit | null {
  if (!(t instanceof Element)) return null;
  const el = t.closest(`[${CONV_ATTR}]`) ?? t.closest('[role="option"]')?.querySelector(`[${CONV_ATTR}]`);
  const convId = el?.getAttribute(CONV_ATTR);
  return el && convId ? { row: rowOf(el), convId } : null;
}

/**
 * The subject text element in a row. The row's aria-label reads "sender subject time preview…".
 * The sender is a span with a title (its email address); the subject is the first untitled
 * text span whose text appears in the aria-label (the preview comes later in the label).
 */
export function findSubjectEl(row: HTMLElement): HTMLElement | null {
  const label = row.getAttribute('aria-label') ?? '';
  let best: HTMLElement | null = null;
  let bestIdx = Infinity;
  for (const s of row.querySelectorAll<HTMLElement>('span')) {
    if (s.children.length || s.hasAttribute('title') || s.closest(`button,[role="checkbox"],${OUR}`)) continue;
    const t = s.textContent?.trim() ?? '';
    if (t.length < 2) continue;
    const i = label.indexOf(t);
    if (i > 0 && i < bestIdx) {
      best = s;
      bestIdx = i;
    }
  }
  return best;
}

export function subjectTextOfRow(row: HTMLElement): string {
  return findSubjectEl(row)?.textContent?.trim() ?? '';
}

/** Conversation id of the single selected row, or null for none / multi-select. */
export function selectedConvId(): string | null {
  const ids = new Set<string>();
  for (const el of document.querySelectorAll('[role="option"][aria-selected="true"]')) {
    const id = el.getAttribute(CONV_ATTR) ?? el.querySelector(`[${CONV_ATTR}]`)?.getAttribute(CONV_ATTR);
    if (id) ids.add(id);
  }
  return ids.size === 1 ? [...ids][0] : null;
}

/** True when Outlook's URL points at an open item (…/id/<itemId>). */
export function urlHasOpenItem(): boolean {
  return /\/id\//.test(location.pathname);
}

const READING_PANE = ['#ReadingPaneContainerId', '#ConversationReadingPaneContainer', '[role="main"][aria-label="Reading Pane"]'];
const CONVERSATION = '#ConversationReadingPaneContainer';

export function findReadingPane(): HTMLElement | null {
  for (const sel of READING_PANE) {
    const el = document.querySelector<HTMLElement>(sel);
    if (el) return el;
  }
  return null;
}

/** The subject heading inside the reading pane (id like "CONV_…_SUBJECT"). */
export function findReadingSubject(pane: HTMLElement): HTMLElement | null {
  return (
    pane.querySelector<HTMLElement>('[role="heading"][id$="_SUBJECT"]') ??
    pane.querySelector<HTMLElement>('[role="heading"]')
  );
}

/**
 * The element to put the banner after: the subject header block, i.e. the heading's ancestor
 * that is a direct child of the conversation container. (The heading's own parent sits in a
 * flex row next to Outlook's action buttons, so inserting there would go sideways.)
 */
export function findBannerAnchor(pane: HTMLElement): HTMLElement | null {
  const heading = findReadingSubject(pane);
  if (!heading) return null;
  const conv = heading.closest<HTMLElement>(CONVERSATION);
  if (conv) {
    let e: HTMLElement = heading;
    while (e.parentElement && e.parentElement !== conv) e = e.parentElement;
    return e;
  }
  return heading.parentElement;
}

/** Outlook's own right-click menu (Fluent UI). */
export const MENU_SELECTOR = '[role="menu"]';
