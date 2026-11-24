/**
 * Add-in task pane dev server: `bun run dev:addin`.
 *
 * Serves dist-addin/ and keeps it rebuilt on save. Open /taskpane.html?mock to try the pane in a
 * browser with a fake Outlook (a bar at the bottom switches between mock emails and shows the
 * categories the pane put on each). For development only; the installed add-in never needs it.
 */
import { join, normalize } from 'node:path';

const PORT = Number(process.env.PORT ?? 8792);
const ROOT = join(import.meta.dir, '..');
const DIST = join(ROOT, 'dist-addin');
const noStore = { 'cache-control': 'no-store' };

const watcher = Bun.spawn(['bun', 'build-addin.mjs', '--watch'], { cwd: ROOT, stdout: 'inherit', stderr: 'inherit' });
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
    const rel = normalize(pathname === '/' ? '/taskpane.html' : pathname);
    if (rel.includes('..')) return new Response('Not found', { status: 404 });
    const file = Bun.file(join(DIST, rel));
    if (await file.exists()) return new Response(file, { headers: noStore });
    return new Response(`${pathname} isn't built yet, or doesn't exist.`, { status: 404 });
  },
});

console.log(`Add-in pane (mock Outlook): http://localhost:${server.port}/taskpane.html?mock`);
