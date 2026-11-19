// Builds each extension entry as a self-contained IIFE bundle (content scripts can't be ES modules).
import { rm } from 'node:fs/promises';
import { build } from 'vite';

const watch = process.argv.includes('--watch');
const entries = [
  ['content', 'src/content/index.ts'],
  ['background', 'src/background.ts'],
  ['editor', 'src/pages/editor.ts'],
  ['options', 'src/pages/options.ts'],
];

// Clear dist once up front. Per-build emptyOutDir would wipe the other bundles on every
// watch-mode rebuild, since all four builds share dist/.
await rm('dist', { recursive: true, force: true });

for (const [i, [name, entry]] of entries.entries()) {
  await build({
    configFile: false,
    publicDir: i === 0 ? 'public' : false,
    logLevel: watch ? 'info' : 'warn',
    clearScreen: false,
    build: {
      outDir: 'dist',
      emptyOutDir: false,
      minify: false,
      target: 'chrome120',
      watch: watch ? {} : null,
      lib: { entry, name: name.replace(/\W/g, '_'), formats: ['iife'], fileName: () => `${name}.js` },
    },
  });
}
console.log(watch ? 'Watching for changes… (reload the extension in chrome://extensions after a rebuild)' : 'Built to dist/');
