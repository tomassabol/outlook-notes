# Outlook Notes & Labels

A Chrome extension for Outlook on the web that adds private notes and emoji labels to email conversations, plus an [Outlook add-in](#outlook-desktop-add-in-new-outlook-for-windows-and-mac) for new Outlook on Windows and Mac.

- **Notes:** plain text, one per conversation. A note icon marks noted emails in the list.
- **B number & work order:** two optional fields per conversation (e.g. `B12345`, `WO-12345`), shown as boxed values only when filled in.
- **Hover card:** rest the pointer on a noted email for about 0.4 s. It shows labels, B number, work order, then the note. Click a label or box to copy its text (labels copy without the emoji); click anywhere else on the card to edit. Text can be drag-selected and copied.
- **Banner:** an open email shows its note and labels in a yellow strip under the subject. It can be collapsed.
- **Labels:** a reusable set of labels (emoji + text + one of Outlook's category colors), shown as category-style tags next to the subject.
- **Editing:** the buttons on the note bar, or right-click an email in the list for **Add note…** / **Labels…**.
- **Options page:** click the toolbar icon to manage labels, browse or search all notes, check storage use, and export or import JSON.

Works on `outlook.office.com`, `outlook.office365.com`, `outlook.cloud.microsoft` (work/school) and `outlook.live.com` (personal).

## Install (developer mode)

```bash
bun install
bun run build
```

1. Open `chrome://extensions`.
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and pick the `dist/` folder.
4. Reload any open Outlook tabs.

While developing, `bun run watch` rebuilds on save. After a rebuild, click the reload icon on the extension card, then reload Outlook.

## Mock Outlook (local dev)

```bash
bun run dev
```

This serves a fake Outlook inbox at http://localhost:8791 (`/?dark` for the dark theme, `/options.html` for the options page). It loads the real built extension from `dist/` with an in-memory `chrome.storage` stub kept in `localStorage`, and rebuilds on save, so you can try changes without reloading the extension. The mock lives in [`dev/mock/`](dev/mock) and the server in [`dev/server.ts`](dev/server.ts). Its DOM mirrors the Outlook attributes the extension relies on (see *How it finds emails*).

## Outlook desktop add-in (new Outlook for Windows and Mac)

New Outlook can't run Chrome extensions, so there is also an **Office add-in** with the same data model. It shows a side pane for the selected email:

- **Note, B number and work order** for the conversation. Changes save automatically.
- **Labels**, as toggles. A label is also applied to the email as an **Outlook category** named after it (e.g. `🔥 Akut`), so it shows in the message list on every device. Outlook only has 25 category colors, so the category uses the closest one to the label's color.
- **Export / Import** of a JSON backup, in the same format as the Chrome extension's options page.

### Where the data lives

- Notes, B numbers, work orders, labels and which conversation has which label are kept in the add-in's browser storage (`localStorage`) **on this computer**. Each computer and each Outlook installation keeps its own copy. Nothing is sent to the add-in's host.
- The only thing written to your mailbox is the **category**: its name and color in your category list, and the category on each labeled email.
- Labels are the source of truth in the pane. When you open an email, its categories are updated to match: missing ones are added, and categories named after a label you removed are taken off. Other categories are left alone. Change labels in the pane rather than with Outlook's own *Categorize* menu.
- Resetting Outlook or clearing its web cache can delete the add-in's storage. Export a backup now and then.
- The Chrome extension and the add-in don't share data.

### How it's served

Office add-ins must load their code from an HTTPS URL. `manifest.xml` points to **GitHub Pages**, which serves only the static files (`taskpane.html`, `taskpane.js`, icons). The [Deploy Outlook add-in](.github/workflows/addin-pages.yml) workflow builds and publishes them on every push to `master`. There's no server and no backend.

One-time setup: in the GitHub repo, open **Settings → Pages** and set **Source** to **GitHub Actions**.

### Install

1. Download `https://tomassabol.github.io/outlook-notes/manifest.xml`.
2. In new Outlook, open **Apps** (or **Get add-ins**) → **My add-ins** → **Custom add-ins** → **Add a custom add-in** → **Add from file…**, and choose `manifest.xml`.
3. Open an email and choose **Notes & Labels** in the message toolbar (in new Outlook it can be under **Apps** or the **…** menu).
4. Click the **pin** icon at the top of the pane. While pinned, it stays open and follows the email you select.

Updates need no reinstall: Outlook loads the latest code from GitHub Pages. Reinstall only when `manifest.xml` changes.

### Develop

```bash
bun run build:addin   # builds dist-addin/ (set ADDIN_BASE_URL to serve it from somewhere else)
bun run dev:addin     # http://localhost:8792/taskpane.html?mock
```

The mock mode runs the pane in a normal browser with fake emails. A bar at the bottom switches emails and shows the categories the pane applied. The code is in [`src/addin/`](src/addin): `taskpane.ts` (UI), `host.ts` (Office.js calls and the mock), `local-area.ts` (storage). The manifest template is [`addin/manifest.template.xml`](addin/manifest.template.xml).

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
src/addin/         Outlook add-in task pane (new Outlook for Windows and Mac)
addin/             add-in manifest template, task pane HTML, icons
build.mjs          builds each entry as a standalone bundle into dist/
build-addin.mjs    builds the add-in into dist-addin/ and writes manifest.xml
```

The note editor is an extension page shown in an iframe overlay. That way keystrokes in the note never reach Outlook, and typing "r" or pressing Delete can't trigger Outlook's shortcuts.
