import { decorateRows } from './rows';
import { initHover } from './hover';
import { initBanner, updateBanner } from './banner';
import { initMenu } from './menu';
import { initEditorFrame } from './editor-frame';
import { initTargetTracking } from './target';
import { loadState, state, watchState } from './state';
import { detectTheme } from './theme';

const MIN_INTERVAL = 120;
let pending = false;
let lastRun = 0;

function schedule(): void {
  if (pending) return;
  pending = true;
  const wait = Math.max(0, MIN_INTERVAL - (performance.now() - lastRun));
  window.setTimeout(() => requestAnimationFrame(run), wait);
}

function run(): void {
  pending = false;
  lastRun = performance.now();
  // After the extension is reloaded, this orphaned script loses chrome.* access. Stop quietly.
  if (!chrome.runtime?.id) return observer.disconnect();
  try {
    state.theme = detectTheme();
    decorateRows(schedule);
    updateBanner();
  } catch (e) {
    console.warn('[outlook-notes]', e);
  }
}

const observer = new MutationObserver(schedule);

async function boot(): Promise<void> {
  await Promise.all([loadState(), initBanner()]);
  initTargetTracking();
  initHover();
  initMenu();
  initEditorFrame();
  watchState(schedule);
  observer.observe(document.body, {
    childList: true,
    subtree: true,
    attributes: true,
    attributeFilter: ['aria-selected', 'data-convid'],
  });
  // Theme switches restyle the top-level elements; watch them without a subtree to stay cheap.
  const themeObserver = new MutationObserver(schedule);
  for (const el of [document.documentElement, document.body]) {
    themeObserver.observe(el, { attributes: true, attributeFilter: ['class', 'style', 'data-theme'] });
  }
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', schedule);
  addEventListener('popstate', schedule);
  schedule();
}

void boot();
