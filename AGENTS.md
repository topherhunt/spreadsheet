# Working on Sheeyit

Electron app: `src/main` (window, menus, file dialogs, IPC; bundled by esbuild to `dist/main/*.cjs`), `src/renderer` (UI; Vite, root `src/renderer`, output `dist/renderer`), `src/core` (the spreadsheet model, with no DOM or Electron, so Vitest runs it directly), `src/shared/api.ts` (the IPC contract and `MenuCommand` list).

## Core model

- `core/model.ts` is the workbook shape. Cells hold the value, an optional formula `f` (stored as in the file, without `=`), style index `s` into `wb.styles`, rich-text runs and a link.
- `core/doc.ts` (`Doc`) is the only way the UI mutates a workbook. Every edit records one undo step: cell edits via `putCells`, layout and structure changes via `snapshotStep`. It emits change events the renderer listens to.
- `core/formula/engine.ts` mirrors the model into HyperFormula, which supplies values only. The model stays the source of truth for formulas. Formulas HyperFormula can't evaluate are `frozen` and keep the file's cached value. Extra functions are in `formula/functions.ts`.
- `core/formula/refs.ts` tokenizes formulas just enough to move references, and leaves every other byte alone. All reference rewriting (copy, fill, insert/delete, rename, cut/paste) goes through `rewriteRefs`.
- `core/xlsx/read.ts` and `write.ts`: untouched sheets (`pristine`) and every part Sheeyit doesn't model are written back byte-identical. Unsupported features are recorded in `features` so the banner can list them. XML with a DOCTYPE is refused.
- `core/csv.ts` keeps the raw text of cells it parsed as numbers, dates or booleans (`raw`), so untouched cells save exactly as read (`007`, `1200.50`).

## Renderer

- `renderer/app.ts` (`App`) is the controller: commands, saving and format prompts, clipboard, context menus, and keyboard handling outside the editor.
- `grid/grid.ts` draws pooled cell divs in four frozen/scrolling panes. Each cell div carries `data-rc="row:col"`. Overlays sit under a transparent `.grid-spacer` scroll layer, and clicks are hit-tested by coordinates.
- `ui/editor.ts` is the in-cell/formula-bar editor. Typing into a cell starts *quick* mode, where arrows commit and move, as in Sheets. Enter or F2 starts full edit mode, where arrows move the caret.
- CSS is utility-first (`styles.css`). Icons are vendored Bootstrap Icons.

## Tests

- `npm test` builds, then runs Vitest (`tests/unit`) and Playwright (`tests/e2e`).
- E2E tests drive the real app through `tests/e2e/harness.ts`:
  - Native dialogs are stubbed in main and answered from queues (`answerSave`, `answerMessage`). An unexpected dialog fails the test.
  - Menu commands are sent straight to window 0, because headless windows are never focused.
  - Apps are killed with SIGKILL, because a graceful quit would ask about unsaved changes.
  - Grid overlays (filter buttons) need `click({ force: true })`, and header context menus need `headerMenu`, because the scroll layer sits on top.
- `clipboard.spec.ts` uses the real system clipboard, so its tests run in order in one worker.
- `packaged-app.spec.ts` runs only after `npm run release`. The bundle's fuses disable `--inspect`, so it attaches over Chromium's debugging port.
- Selectors are `qa-*` classes, not visible text.
