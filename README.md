# Outlook Notes & Labels

A Chrome extension for Outlook on the web that adds private notes and emoji labels to email conversations.

- **Notes:** plain text, one per conversation. A note icon marks noted emails in the list.
- **Hover card:** rest the pointer on a noted email in the list for about 0.4 s to read its note and labels.
- **Banner:** an open email shows its note and labels in a yellow strip under the subject. It can be collapsed.
- **Labels:** a reusable set of labels (emoji + text + one of Outlook's category colors), shown as category-style tags next to the subject.
- **Editing:** the buttons on the note bar, or right-click an email in the list for **Add note…** / **Labels…**.
- **Options page:** click the toolbar icon to manage labels, browse or search all notes, check storage use, and export or import JSON.

Works on `outlook.office.com`, `outlook.office365.com`, `outlook.cloud.microsoft` (work/school) and `outlook.live.com` (personal).

## Install (developer mode)

```bash
npm install
npm run build
```

1. Open `chrome://extensions`.
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and pick the `dist/` folder.
4. Reload any open Outlook tabs.

While developing, `npm run watch` rebuilds on save. After a rebuild, click the reload icon on the extension card, then reload Outlook.

## Look and feel

The UI is built to blend into Outlook:

- Styles use Outlook's own Fluent UI v9 design tokens (`--colorNeutralForeground1`, `--shadow16`, …), read live from the page. That means light, dark and custom Outlook themes are matched automatically. The alias layer is in [`src/shared/tokens.css`](src/shared/tokens.css).
- Label colors are Outlook's category palette (`--colorPalette<Name>Background2/Foreground2`).
- Icons are Microsoft's [Fluent System Icons](https://github.com/microsoft/fluentui-system-icons), the set Outlook uses.
- Components follow their Fluent counterparts: the note bar is a MessageBar, the hover card is a Popover, the editor is a Dialog, and the right-click items are MenuItems.
- The editor iframe can't inherit Outlook's CSS variables, so the content script sends it the current token values when it opens.

## Data & privacy

- Everything is stored in `chrome.storage.sync`, so it follows your Chrome profile. Nothing is sent anywhere else, and nothing is written to your mailbox.
- Notes don't appear in Outlook mobile or desktop apps.
- Chrome sync allows about **100 KB** in total, which is enough for a few hundred notes. The options page shows how much you've used. Export a backup now and then.
- Conversation IDs are stored as short hashes (`c:<hash>`), each with a 60-character subject snapshot for the options page.

## How it finds emails

Outlook on the web uses generated CSS class names, so the extension relies only on stable attributes. They're all in [`src/content/selectors.ts`](src/content/selectors.ts):

| What | How |
|---|---|
| List rows | `div[role="option"][data-convid]`, the conversation ID |
| Subject in a row | first untitled text span whose text appears in the row's `aria-label` |
| Open conversation | the single `aria-selected="true"` row, plus a `…/id/…` URL fallback |
| Reading pane | `#ReadingPaneContainerId`; subject heading `[role="heading"][id$="_SUBJECT"]` |
| Banner position | after the subject header block inside `#ConversationReadingPaneContainer` |
| Right-click menu | Outlook's `[role="menu"]`; our items are prepended to it |

If Microsoft changes the DOM and pills or the banner disappear, that file is where to fix it.

## Project layout

```
public/            manifest.json, editor.html, options.html
src/background.ts  seeds default labels, Chrome context-menu fallback, toolbar → options
src/content/       content script: rows, hover card, banner, menu injection, theme
src/pages/         editor (runs in an iframe) and options page
src/shared/        storage, types, DOM helpers, shared CSS tokens
build.mjs          builds each entry as a standalone bundle into dist/
```

The note editor is an extension page shown in an iframe overlay. That way keystrokes in the note never reach Outlook, and typing "r" or pressing Delete can't trigger Outlook's shortcuts.
