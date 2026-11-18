import { DEFAULT_LABELS, getLabels, migrateLabels, saveLabels } from './shared/store';
import type { EditorFocus, Msg } from './shared/types';

const OUTLOOK = [
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
  'https://outlook.cloud.microsoft/*',
  'https://outlook.live.com/*',
];

chrome.runtime.onInstalled.addListener(async ({ reason }) => {
  if (reason === 'install' && !(await getLabels()).length) await saveLabels(DEFAULT_LABELS);
  if (reason === 'update') await migrateLabels();

  chrome.contextMenus.removeAll(() => {
    const common = { contexts: ['page', 'selection', 'link'] as ['page', 'selection', 'link'], documentUrlPatterns: OUTLOOK };
    chrome.contextMenus.create({ id: 'onl-note', title: 'Add or edit note…', ...common });
    chrome.contextMenus.create({ id: 'onl-labels', title: 'Labels…', ...common });
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (!tab?.id) return;
  const focus: EditorFocus = info.menuItemId === 'onl-labels' ? 'labels' : 'note';
  const msg: Msg = { type: 'open-editor', focus };
  chrome.tabs.sendMessage(tab.id, msg, { frameId: info.frameId ?? 0 }).catch(() => {
    /* content script not ready (e.g. tab opened before install) */
  });
});

chrome.action.onClicked.addListener(() => chrome.runtime.openOptionsPage());
