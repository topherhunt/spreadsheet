# Wish list

## Windows release

The NSIS target in `package.json` and the **Build** and **Windows Tests** workflows exist, but neither workflow has run yet. To ship:

- Run both workflows and fix whatever fails on Windows, then smoke-test the `.exe` on a real machine.
- Attach the `.exe` from the Build workflow to each release, since `scripts/release.sh` only uploads the `.dmg`.
- Add Windows install steps to the README. The installer is unsigned, so SmartScreen blocks it until you click **More info > Run anyway**.
- Mention Windows again in the README's opening line.

## Comments

Let people view, add, edit, resolve and delete comments. The app doesn't show comments yet: they are kept in the file unmodeled, listed in the banner ("Notes/comments"), and moved with row/column inserts and deletes by `core/xlsx/follow.ts`.

- Excel has two kinds. Notes (`comments1.xml`, drawn by a VML part) are plain text boxes. Threaded comments (`threadedComments`, with a `persons` part) have replies and can be resolved. Excel writes a note alongside each threaded comment so older readers still see it, so the two must stay in sync.
- Model comments on the cell, so that sort, cut/paste and fill carry them along, which `follow.ts` doesn't do.
- Show a corner marker on commented cells and a popover on hover or click, plus Add/Edit/Delete/Resolve commands in the cell context menu and a list of every comment in the workbook.
- CSV can't store comments, so `delimitedLosses` should warn about them.
