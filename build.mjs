// Builds each extension entry as a self-contained IIFE bundle (content scripts can't be ES modules).
import { build } from 'vite';

const watch = process.argv.includes('--watch');
const entries = [
  ['content', 'src/content/index.ts'],
  ['background', 'src/background.ts'],
  ['editor', 'src/pages/editor.ts'],
  ['options', 'src/pages/options.ts'],
];

for (const [i, [name, entry]] of entries.entries()) {
  await build({
    configFile: false,
    publicDir: i === 0 ? 'public' : false,
    logLevel: 'warn',
    build: {
      outDir: 'dist',
      emptyOutDir: i === 0,
      minify: false,
      target: 'chrome120',
      watch: watch ? {} : null,
      lib: { entry, name: name.replace(/\W/g, '_'), formats: ['iife'], fileName: () => `${name}.js` },
    },
  });
}
console.log(watch ? 'Watching for changes…' : 'Built to dist/');
