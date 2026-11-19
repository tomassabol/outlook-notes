/**
 * Mock Outlook dev server: `bun run dev`.
 *
 * Serves a fake Outlook inbox (dev/mock) that loads the real built extension from dist/ with a
 * chrome.storage stub, and keeps dist/ rebuilt on save. Unknown paths (e.g. /mail/inbox/id/…,
 * which the mock pushes like Outlook does) fall back to the inbox page, so reloads work.
 */
import { join } from 'node:path';

const PORT = Number(process.env.PORT ?? 8791);
const ROOT = join(import.meta.dir, '..');
const MOCK = join(ROOT, 'dev/mock');
const DIST = join(ROOT, 'dist');

const MOCK_FILES = new Set(['/index.html', '/options.html', '/editor.html', '/chrome-stub.js']);
const DIST_FILE = /^\/(?:[\w.-]+\.(?:js|json)|icons\/[\w.-]+\.png)$/;
const noStore = { 'cache-control': 'no-store' };

const watcher = Bun.spawn(['bun', 'build.mjs', '--watch'], { cwd: ROOT, stdout: 'inherit', stderr: 'inherit' });
for (const sig of ['SIGINT', 'SIGTERM'] as const) {
  process.on(sig, () => {
    watcher.kill();
    process.exit(0);
  });
}

const server = Bun.serve({
  port: PORT,
  async fetch(req) {
    const { pathname } = new URL(req.url);
    if (MOCK_FILES.has(pathname)) return new Response(Bun.file(join(MOCK, pathname)), { headers: noStore });
    if (DIST_FILE.test(pathname)) {
      const file = Bun.file(join(DIST, pathname));
      if (await file.exists()) return new Response(file, { headers: noStore });
      return new Response(`${pathname} isn't built yet. The watcher is still building; reload in a moment.`, { status: 404 });
    }
    if (pathname === '/favicon.ico') return new Response(null, { status: 204 });
    // Everything else is a mock Outlook route.
    return new Response(Bun.file(join(MOCK, 'index.html')), { headers: { ...noStore, 'content-type': 'text/html; charset=utf-8' } });
  },
});

console.log(`Mock Outlook: http://localhost:${server.port}/  (dark theme: /?dark, options page: /options.html)`);
