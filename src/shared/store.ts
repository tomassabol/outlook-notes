import { PALETTE, paletteName } from './palette';
import type { ConvData, Label } from './types';

export const LABELS_KEY = 'labels';
export const CONV_PREFIX = 'c:';
/** The subset of `chrome.storage.StorageArea` the store uses, so other hosts can plug in their own. */
export interface KVArea {
  get(keys: string | null): Promise<Record<string, unknown>>;
  set(items: Record<string, unknown>): Promise<void>;
  remove(keys: string | string[]): Promise<void>;
  getBytesInUse(keys: null): Promise<number>;
  readonly QUOTA_BYTES: number;
}

let customArea: KVArea | null = null;
/** Use a different backing store (the Outlook add-in uses localStorage). Call before any other store function. */
export function setStorageArea(a: KVArea): void {
  customArea = a;
}
const area = (): KVArea => customArea ?? (chrome.storage.sync as unknown as KVArea);

/** Fast 53-bit string hash. Conversation IDs are ~150 chars; hashing keeps keys short. */
function cyrb53(str: string, seed = 0): number {
  let h1 = 0xdeadbeef ^ seed;
  let h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return 4294967296 * (2097151 & h2) + (h1 >>> 0);
}

const keyCache = new Map<string, string>();
export function convKey(convId: string): string {
  let k = keyCache.get(convId);
  if (!k) {
    k = CONV_PREFIX + cyrb53(convId).toString(36);
    keyCache.set(convId, k);
  }
  return k;
}

export function hasContent(d: ConvData | undefined | null): d is ConvData {
  return !!d && (!!d.n?.trim() || !!d.l?.length || !!d.b?.trim() || !!d.w?.trim());
}

export const DEFAULT_LABELS: Label[] = [
  { id: 'urgent', emoji: '🔥', text: 'Akut', color: 'Red' },
  { id: 'waiting', emoji: '⏳', text: 'Inväntar svar', color: 'Marigold' },
  { id: 'followup', emoji: '📌', text: 'Följ upp', color: 'Blue' },
  { id: 'invoice', emoji: '💰', text: 'Offert', color: 'Green' },
];

/** Default labels renamed to Swedish (v0.3.1, Följ upp in v0.3.3); existing labels still using the old name are updated. */
const RENAMED_DEFAULTS: Record<string, [from: string, to: string]> = {
  urgent: ['Urgent', 'Akut'],
  waiting: ['Waiting', 'Inväntar svar'],
  invoice: ['Invoice', 'Offert'],
  followup: ['Follow up', 'Följ upp'],
};

export function newId(): string {
  return crypto.randomUUID().replace(/-/g, '').slice(0, 8);
}

export async function getLabels(): Promise<Label[]> {
  const r = await area().get(LABELS_KEY);
  return Array.isArray(r[LABELS_KEY]) ? (r[LABELS_KEY] as Label[]) : [];
}

export async function saveLabels(labels: Label[]): Promise<void> {
  await area().set({ [LABELS_KEY]: labels });
}

export async function getConv(key: string): Promise<ConvData | undefined> {
  const r = await area().get(key);
  return r[key] as ConvData | undefined;
}

export async function getAllConvs(): Promise<Record<string, ConvData>> {
  const all = await area().get(null);
  const out: Record<string, ConvData> = {};
  for (const [k, v] of Object.entries(all)) if (k.startsWith(CONV_PREFIX)) out[k] = v as ConvData;
  return out;
}

/** Saves, or removes the entry when it has neither note nor labels. */
export async function saveConv(key: string, data: ConvData | null): Promise<void> {
  if (hasContent(data)) await area().set({ [key]: data });
  else await area().remove(key);
}

/** Removes a label from the library and from every conversation using it. */
export async function deleteLabel(id: string): Promise<void> {
  const all = await area().get(null);
  const labels = ((all[LABELS_KEY] as Label[] | undefined) ?? []).filter((l) => l.id !== id);
  const updates: Record<string, unknown> = { [LABELS_KEY]: labels };
  const removals: string[] = [];
  for (const [k, v] of Object.entries(all)) {
    if (!k.startsWith(CONV_PREFIX)) continue;
    const d = v as ConvData;
    if (!d.l?.includes(id)) continue;
    const next: ConvData = { ...d, l: d.l.filter((x) => x !== id) };
    if (!next.l!.length) delete next.l;
    if (hasContent(next)) updates[k] = next;
    else removals.push(k);
  }
  await area().set(updates);
  if (removals.length) await area().remove(removals);
}

export async function usage(): Promise<{ used: number; quota: number; items: number }> {
  const [used, all] = await Promise.all([area().getBytesInUse(null), area().get(null)]);
  return { used, quota: area().QUOTA_BYTES, items: Object.keys(all).length };
}

export function friendlyError(e: unknown): string {
  const msg = e instanceof Error ? e.message : String(e);
  if (/QUOTA_BYTES|quota/i.test(msg)) {
    return 'Chrome sync storage is full. Delete old notes on the extension’s options page, or shorten this note.';
  }
  if (/MAX_WRITE_OPERATIONS/i.test(msg)) return 'Too many saves in a short time. Wait a minute and try again.';
  return msg;
}

/**
 * Converts labels saved with v0.1 hex colors to palette names, and renames default labels the
 * user never renamed (English → Swedish). Returns true if anything changed.
 */
export async function migrateLabels(): Promise<boolean> {
  const labels = await getLabels();
  let changed = false;
  for (const l of labels) {
    const rename = RENAMED_DEFAULTS[l.id];
    if (rename && l.text === rename[0]) {
      l.text = rename[1];
      changed = true;
    }
    if (!(l.color in PALETTE)) {
      l.color = paletteName(l.color);
      changed = true;
    }
  }
  if (changed) await saveLabels(labels);
  return changed;
}

export function isLabel(x: unknown): x is Label {
  const l = x as Label;
  return !!l && typeof l.id === 'string' && typeof l.emoji === 'string' && typeof l.text === 'string' &&
    typeof l.color === 'string' && (l.color in PALETTE || /^#[0-9a-f]{6}$/i.test(l.color));
}

export function isConvData(x: unknown): x is ConvData {
  const d = x as ConvData;
  return !!d && typeof d === 'object' && typeof d.u === 'number' &&
    (d.n === undefined || typeof d.n === 'string') &&
    (d.b === undefined || typeof d.b === 'string') &&
    (d.w === undefined || typeof d.w === 'string') &&
    (d.l === undefined || (Array.isArray(d.l) && d.l.every((i) => typeof i === 'string')));
}
