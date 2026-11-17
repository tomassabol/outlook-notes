import { findReadingPane, findReadingSubject, findRows, rowFromTarget, selectedConvId, subjectTextOfRow, urlHasOpenItem } from './selectors';
import { dataFor } from './state';

/** Last conversation the user clicked, for when no list row is selected (e.g. full-page reading view). */
let lastOpened: string | null = null;

export function initTargetTracking(): void {
  document.addEventListener(
    'click',
    (e) => {
      const hit = rowFromTarget(e.target);
      if (hit) lastOpened = hit.convId;
    },
    true,
  );
}

/** Conversation currently shown in the reading pane, if we can tell. */
export function currentOpenConvId(): string | null {
  const sel = selectedConvId();
  if (sel) return sel;
  return urlHasOpenItem() ? lastOpened : null;
}

export function subjectFor(convId: string): string {
  const row = findRows().find((r) => r.convId === convId)?.row;
  const fromRow = row ? subjectTextOfRow(row) : '';
  if (fromRow) return fromRow;
  const pane = findReadingPane();
  const fromPane = pane ? findReadingSubject(pane)?.textContent?.trim() : '';
  return fromPane || dataFor(convId)?.s || '';
}
