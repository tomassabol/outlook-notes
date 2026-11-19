(() => {
  const listeners = [];
  const mk = (name, data) => ({
    QUOTA_BYTES: 102400,
    async get(keys) {
      if (keys == null) return structuredClone(data);
      const ks = typeof keys === 'string' ? [keys] : Array.isArray(keys) ? keys : Object.keys(keys);
      const o = {}; for (const k of ks) if (k in data) o[k] = structuredClone(data[k]); return o;
    },
    async set(obj) {
      const ch = {};
      for (const [k, v] of Object.entries(obj)) { ch[k] = { oldValue: data[k], newValue: structuredClone(v) }; data[k] = structuredClone(v); }
      localStorage.setItem('stub-' + name, JSON.stringify(data));
      listeners.forEach((f) => f(ch, name));
    },
    async remove(keys) {
      const ch = {};
      for (const k of [].concat(keys)) { ch[k] = { oldValue: data[k] }; delete data[k]; }
      localStorage.setItem('stub-' + name, JSON.stringify(data));
      listeners.forEach((f) => f(ch, name));
    },
    async clear() { await this.remove(Object.keys(data)); },
    async getBytesInUse() { return Object.entries(data).reduce((n, [k, v]) => n + k.length + JSON.stringify(v).length, 0); },
  });
  const load = (n) => JSON.parse(localStorage.getItem('stub-' + n) || 'null');
  const sync = load('sync') ?? {
    labels: [
      // Same as DEFAULT_LABELS in src/shared/store.ts
      { id: 'urgent', emoji: '🔥', text: 'Akut', color: 'Red' },
      { id: 'waiting', emoji: '⏳', text: 'Inväntar svar', color: 'Marigold' },
      { id: 'followup', emoji: '📌', text: 'Följ upp', color: 'Blue' },
      { id: 'invoice', emoji: '💰', text: 'Offert', color: 'Green' },
    ],
  };
  const stub = {
    runtime: { id: 'stub', getURL: (p) => new URL('/' + p, location.origin).href, onMessage: { addListener() {} }, openOptionsPage() {} },
    storage: { sync: mk('sync', sync), local: mk('local', load('local') ?? {}), onChanged: { addListener: (f) => listeners.push(f) } },
  };
  Object.defineProperty(window, 'chrome', { value: stub, configurable: true });
})();
