// Builds the Outlook add-in (task pane) into dist-addin/, ready to publish as static files.
//
//   ADDIN_BASE_URL=https://example.github.io/outlook-notes/ bun build-addin.mjs
//
// ADDIN_BASE_URL is where dist-addin/ will be served from (HTTPS, trailing slash). It is baked
// into dist-addin/manifest.xml, the file you install in Outlook.
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { build } from 'vite';

const watch = process.argv.includes('--watch');
const base = (process.env.ADDIN_BASE_URL || 'https://tomassabol.github.io/outlook-notes/').replace(/\/?$/, '/');
const out = 'dist-addin';
const pkg = JSON.parse(await readFile('package.json', 'utf8'));

await rm(out, { recursive: true, force: true });
const result = await build({
  configFile: false,
  publicDir: 'addin',
  logLevel: watch ? 'info' : 'warn',
  clearScreen: false,
  build: {
    outDir: out,
    emptyOutDir: false,
    minify: false,
    target: ['chrome110', 'safari16'],
    watch: watch ? {} : null,
    lib: { entry: 'src/addin/taskpane.ts', name: 'taskpane', formats: ['iife'], fileName: () => 'taskpane.js' },
  },
});

// Manifest: fill in URLs, and give the outer (v1.0) overrides the same commands minus pinning,
// which only v1.1 supports. Clients that understand v1.1 use the inner block.
async function writeManifest() {
  let xml = await readFile('addin/manifest.template.xml', 'utf8');
  xml = xml
    .replaceAll('{{BASE}}', base)
    .replaceAll('{{ORIGIN}}', new URL(base).origin)
    .replaceAll('{{VERSION}}', `${pkg.version}.0`);
  const inner = /(<VersionOverrides [^>]*VersionOverridesV1_1">)([\s\S]*?)(\n\s*<\/VersionOverrides>\n\s*<\/VersionOverrides>)/.exec(xml);
  if (!inner) throw new Error('manifest template: nested VersionOverrides not found');
  const v10Body = inner[2].replace(/\s*<SupportsPinning>true<\/SupportsPinning>/, '').replace(/^ {2}/gm, '');
  xml = xml.replace(inner[1], `${v10Body.trimEnd()}\n    ${inner[1]}`);
  await writeFile(`${out}/manifest.xml`, xml);
  await rm(`${out}/manifest.template.xml`, { force: true });
}

if (watch) {
  // In watch mode build() returns before the first build; write the manifest after each one.
  result.on('event', (e) => {
    if (e.code === 'END') void mkdir(out, { recursive: true }).then(writeManifest);
  });
} else {
  await writeManifest();
}

console.log(`Built add-in to ${out}/ (served from ${base})`);
