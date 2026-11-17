import type { ConvData, Label } from './types';

export const LABELS_KEY = 'labels';
export const CONV_PREFIX = 'c:';
const area = () => chrome.storage.sync;

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
  return !!d && (!!d.n?.trim() || !!d.l?.length);
}

export const DEFAULT_LABELS: Label[] = [
  { id: 'urgent', emoji: '🔥', text: 'Urgent', color: '#d13438' },
  { id: 'waiting', emoji: '⏳', text: 'Waiting', color: '#f2c94c' },
  { id: 'followup', emoji: '📌', text: 'Follow up', color: '#0f6cbd' },
  { id: 'invoice', emoji: '💰', text: 'Invoice', color: '#107c41' },
];

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
  return { used, quota: chrome.storage.sync.QUOTA_BYTES, items: Object.keys(all).length };
}

export function friendlyError(e: unknown): string {
  const msg = e instanceof Error ? e.message : String(e);
  if (/QUOTA_BYTES|quota/i.test(msg)) {
    return 'Chrome sync storage is full. Delete old notes on the extension’s options page, or shorten this note.';
  }
  if (/MAX_WRITE_OPERATIONS/i.test(msg)) return 'Too many saves in a short time. Wait a minute and try again.';
  return msg;
}

/** Readable text color (dark or white) for a label background. */
export function textColorFor(hex: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return '#ffffff';
  const n = parseInt(m[1], 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum > 0.4 ? '#1b1a19' : '#ffffff';
}

export function isLabel(x: unknown): x is Label {
  const l = x as Label;
  return !!l && typeof l.id === 'string' && typeof l.emoji === 'string' && typeof l.text === 'string' &&
    typeof l.color === 'string' && /^#[0-9a-f]{6}$/i.test(l.color);
}

export function isConvData(x: unknown): x is ConvData {
  const d = x as ConvData;
  return !!d && typeof d === 'object' && typeof d.u === 'number' &&
    (d.n === undefined || typeof d.n === 'string') &&
    (d.l === undefined || (Array.isArray(d.l) && d.l.every((i) => typeof i === 'string')));
}
