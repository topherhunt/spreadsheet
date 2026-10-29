# Wish list

## Windows release

The NSIS target in `package.json` and the **Build** and **Windows Tests** workflows exist, but neither workflow has run yet. To ship:

- Run both workflows and fix whatever fails on Windows, then smoke-test the `.exe` on a real machine.
- Attach the `.exe` from the Build workflow to each release, since `scripts/release.sh` only uploads the `.dmg`.
- Add Windows install steps to the README. The installer is unsigned, so SmartScreen blocks it until you click **More info > Run anyway**.
- Mention Windows again in the README's opening line.

## Advanced formulas tier 1

Scalar Excel functions HyperFormula lacks, added to `ExtraFunctions` in `core/formula/functions.ts`. Until then, formulas using them are frozen at the file's cached value.

- Bonds and securities: `ACCRINT`, `ACCRINTM`, `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `YIELD`, `YIELDDISC`, `YIELDMAT`, `DURATION`, `MDURATION`, `INTRATE`, `RECEIVED`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `AMORDEGRC`, `AMORLINC`, `VDB`.
- Other: `CONVERT`, `ERF.PRECISE`, `ERFC.PRECISE`, `ECMA.CEILING`, `PERCENTOF`, `PROB`, `BINOM.DIST.RANGE`, `AGGREGATE`, `ARRAYTOTEXT`, `VALUETOTEXT`, `BAHTTEXT`, and the double-byte text functions (`LENB`, `FINDB`, `SEARCHB`, `LEFTB`, `REPLACEB`, `ASC`, `JIS`).
- Take expected values from a file saved by real Excel. The bond functions depend on day-count conventions, and the documented formulas don't pin down every edge case.

## Advanced formulas tier 2

Functions that return arrays: `TOCOL`, `TOROW`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `WRAPROWS`, `WRAPCOLS`, `SORTBY`, `RANDARRAY`, `MUNIT`, `MINVERSE`, `LINEST`, `LOGEST`, `TREND`, `GROWTH`, `FREQUENCY`, `MODE.MULT`, `TEXTSPLIT`.

- HyperFormula fixes an array's size when it parses the formula, before it knows any values. For `FILTER`, the engine allocates the input's size and trims the empty cells (`arraySize` in `engine.ts`). Use the same approach wherever an input range bounds the size.
- `TEXTSPLIT`, and sizes taken from another cell (`TAKE(A:A, B1)`), have no such bound. Allocating too much causes false `#SPILL!` errors next to other data, so those cases may need to stay frozen.

## Comments

Let people view, add, edit, resolve and delete comments. The app doesn't show comments yet: they are kept in the file unmodeled, listed in the banner ("Notes/comments"), and moved with row/column inserts and deletes by `core/xlsx/follow.ts`.

- Excel has two kinds. Notes (`comments1.xml`, drawn by a VML part) are plain text boxes. Threaded comments (`threadedComments`, with a `persons` part) have replies and can be resolved. Excel writes a note alongside each threaded comment so older readers still see it, so the two must stay in sync.
- Model comments on the cell, so that sort, cut/paste and fill carry them along, which `follow.ts` doesn't do.
- Show a corner marker on commented cells and a popover on hover or click, plus Add/Edit/Delete/Resolve commands in the cell context menu and a list of every comment in the workbook.
- CSV can't store comments, so `delimitedLosses` should warn about them.
