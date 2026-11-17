import { CONV_PREFIX, LABELS_KEY, convKey, getAllConvs, getLabels } from '../shared/store';
import type { ConvData, Label, Theme } from '../shared/types';

export const state = {
  labels: [] as Label[],
  /** keyed by storage key (`c:<hash>`) */
  convs: new Map<string, ConvData>(),
  theme: 'light' as Theme,
  /** bumps when the label library changes, so cached renders refresh */
  labelsRev: 0,
};

export async function loadState(): Promise<void> {
  const [labels, convs] = await Promise.all([getLabels(), getAllConvs()]);
  state.labels = labels;
  state.convs = new Map(Object.entries(convs));
}

export function watchState(onUpdate: () => void): void {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'sync') return;
    for (const [k, { newValue }] of Object.entries(changes)) {
      if (k === LABELS_KEY) {
        state.labels = Array.isArray(newValue) ? newValue : [];
        state.labelsRev++;
      } else if (k.startsWith(CONV_PREFIX)) {
        if (newValue) state.convs.set(k, newValue as ConvData);
        else state.convs.delete(k);
      }
    }
    onUpdate();
  });
}

export const dataFor = (convId: string): ConvData | undefined => state.convs.get(convKey(convId));

export function labelsOf(d: ConvData | undefined): Label[] {
  if (!d?.l) return [];
  return d.l.map((id) => state.labels.find((l) => l.id === id)).filter((l): l is Label => !!l);
}

/** Changes whenever what we'd render for this conversation changes. */
export function signature(convId: string, d: ConvData | undefined, extra = ''): string {
  return `${convId}|${d?.u ?? 0}|${state.labelsRev}|${state.theme}|${extra}`;
}
