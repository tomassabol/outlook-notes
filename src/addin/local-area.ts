import type { KVArea } from '../shared/store';

/**
 * `localStorage`-backed storage for the Outlook add-in. Data lives only in this Outlook
 * installation's web view on this computer; nothing is sent to the add-in's host or the mailbox.
 * Each store key is saved as its own JSON value under a prefix.
 */
const PREFIX = 'onl:';

function readAll(): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k?.startsWith(PREFIX)) continue;
    try {
      out[k.slice(PREFIX.length)] = JSON.parse(localStorage.getItem(k) ?? 'null');
    } catch {
      // ignore a corrupt entry rather than failing every read
    }
  }
  return out;
}

export const localArea: KVArea = {
  QUOTA_BYTES: 5 * 1024 * 1024,
  async get(keys) {
    if (keys === null) return readAll();
    const raw = localStorage.getItem(PREFIX + keys);
    return raw === null ? {} : { [keys]: JSON.parse(raw) };
  },
  async set(items) {
    for (const [k, v] of Object.entries(items)) localStorage.setItem(PREFIX + k, JSON.stringify(v));
  },
  async remove(keys) {
    for (const k of Array.isArray(keys) ? keys : [keys]) localStorage.removeItem(PREFIX + k);
  },
  async getBytesInUse() {
    let n = 0;
    for (const [k, v] of Object.entries(readAll())) n += (PREFIX + k).length + JSON.stringify(v).length;
    return n * 2; // UTF-16
  },
};
