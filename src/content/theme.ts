import type { Theme } from '../shared/types';

function bgOf(el: Element | null): [number, number, number] | null {
  for (let e = el; e; e = e.parentElement) {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(getComputedStyle(e).backgroundColor);
    if (m && (m[4] === undefined || parseFloat(m[4]) > 0.5)) return [+m[1], +m[2], +m[3]];
  }
  return null;
}

/**
 * Outlook's theme isn't exposed reliably, so read the luminance of a real surface. Many Outlook
 * containers are transparent; list rows and the list itself have a solid background. Outlook's
 * theme is independent of the OS, so the OS setting is only a last resort.
 */
export function detectTheme(): Theme {
  const probes = [
    document.querySelector('[role="option"][data-convid][aria-selected="false"]'),
    document.querySelector('[role="listbox"]'),
    document.querySelector('#ReadingPaneContainerId'),
    document.querySelector('[role="main"]'),
    document.body,
  ];
  for (const p of probes) {
    const rgb = p ? bgOf(p) : null;
    if (!rgb) continue;
    const [r, g, b] = rgb;
    return 0.299 * r + 0.587 * g + 0.114 * b < 110 ? 'dark' : 'light';
  }
  return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
