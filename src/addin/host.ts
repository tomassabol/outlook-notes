/**
 * Everything the task pane needs from Outlook, behind one interface, so the pane can also run in
 * a plain browser against a mock (`taskpane.html?mock`) for local development.
 */
import { paletteName } from '../shared/palette';
import type { Label } from '../shared/types';

export interface CurrentItem {
  convId: string;
  subject: string;
  /** Identifies the mailbox, so master categories are created once per account. */
  mailbox: string;
}

export interface Host {
  current(): CurrentItem | null;
  onItemChanged(cb: () => void): void;
  itemCategories(): Promise<string[]>;
  addItemCategories(names: string[]): Promise<void>;
  removeItemCategories(names: string[]): Promise<void>;
  masterCategories(): Promise<string[]>;
  addMasterCategories(cats: { name: string; color: string }[]): Promise<void>;
}

/** The Outlook category name for a label, e.g. "🔥 Akut". */
export const categoryName = (l: Label): string => `${l.emoji} ${l.text}`.trim();

/**
 * Label palette name → Outlook category preset. Outlook only has 25 fixed category colors, so
 * this is the closest preset by hue; the pane itself still shows the exact label color.
 */
const PRESET: Record<string, number> = {
  Red: 0, Cranberry: 9, Pumpkin: 1, Peach: 16, Marigold: 3, Gold: 18, Brass: 17, Brown: 2,
  Forest: 6, LightGreen: 4, Green: 19, Seafoam: 5, LightTeal: 5, Teal: 20, Steel: 10,
  Blue: 7, RoyalBlue: 22, Cornflower: 7, Lavender: 8, Purple: 23, Grape: 24, Pink: 9,
  Magenta: 24, Platinum: 12, Anchor: 13,
};
export const categoryColor = (l: Label): string => `Preset${PRESET[paletteName(l.color)] ?? 12}`;

// ---------- Office.js ----------

function call<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error?.message ?? 'Outlook request failed')))),
  );
}

export function officeHost(): Host {
  const mailbox = Office.context.mailbox;
  const item = () => mailbox.item as Office.MessageRead | null | undefined;
  return {
    current() {
      const it = item();
      if (!it?.conversationId) return null;
      return { convId: it.conversationId, subject: it.subject ?? '', mailbox: mailbox.userProfile.emailAddress };
    },
    onItemChanged(cb) {
      mailbox.addHandlerAsync(Office.EventType.ItemChanged, cb);
    },
    async itemCategories() {
      const it = item();
      if (!it?.categories) return [];
      const cats = await call<Office.CategoryDetails[]>((cb) => it.categories.getAsync(cb));
      return (cats ?? []).map((c) => c.displayName);
    },
    async addItemCategories(names) {
      const it = item();
      if (it?.categories && names.length) await call<void>((cb) => it.categories.addAsync(names, cb));
    },
    async removeItemCategories(names) {
      const it = item();
      if (it?.categories && names.length) await call<void>((cb) => it.categories.removeAsync(names, cb));
    },
    async masterCategories() {
      const cats = await call<Office.CategoryDetails[]>((cb) => mailbox.masterCategories.getAsync(cb));
      return (cats ?? []).map((c) => c.displayName);
    },
    async addMasterCategories(cats) {
      if (!cats.length) return;
      const details = cats.map((c) => ({ displayName: c.name, color: Office.MailboxEnums.CategoryColor[c.color as keyof typeof Office.MailboxEnums.CategoryColor] }));
      await call<void>((cb) => mailbox.masterCategories.addAsync(details, cb));
    },
  };
}

// ---------- mock for local development ----------

export function mockHost(): Host {
  const items = [
    { convId: 'AAQkAGMock1', subject: 'Offert för ombyggnad, kvarter 7' },
    { convId: 'AAQkAGMock2', subject: 'Re: Arbetsorder WO-20931 – leverans' },
    { convId: 'AAQkAGMock3', subject: 'Möte på torsdag' },
  ];
  const cats = new Map<string, Set<string>>();
  const master = new Set<string>();
  let idx = 0;
  let listener: (() => void) | null = null;
  const bar = document.createElement('div');
  bar.style.cssText = 'position:fixed;bottom:0;left:0;right:0;padding:6px;background:#333;color:#fff;font:12px sans-serif;display:flex;gap:6px;align-items:center;z-index:9';
  bar.append('Mock:');
  for (const [i, it] of [...items, null].entries()) {
    const b = document.createElement('button');
    b.textContent = it ? `Mail ${i + 1}` : 'None';
    b.onclick = () => {
      idx = i;
      listener?.();
    };
    bar.append(b);
  }
  const catsOut = document.createElement('span');
  bar.append(catsOut);
  const showCats = () => (catsOut.textContent = `categories: ${[...(cats.get(items[idx]?.convId) ?? [])].join(', ') || '—'}`);
  setInterval(showCats, 300);
  document.addEventListener('DOMContentLoaded', () => document.body.append(bar));
  if (document.body) document.body.append(bar);
  const set = () => {
    const id = items[idx]?.convId;
    if (!cats.has(id)) cats.set(id, new Set());
    return cats.get(id)!;
  };
  return {
    current: () => (items[idx] ? { ...items[idx], mailbox: 'mock@example.com' } : null),
    onItemChanged: (cb) => (listener = cb),
    itemCategories: async () => [...set()],
    addItemCategories: async (n) => n.forEach((x) => set().add(x)),
    removeItemCategories: async (n) => n.forEach((x) => set().delete(x)),
    masterCategories: async () => [...master],
    addMasterCategories: async (c) => c.forEach((x) => master.add(x.name)),
  };
}
