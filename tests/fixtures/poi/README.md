# Fixtures from Apache POI

These workbooks come from the Apache POI project's test data (`test-data/spreadsheet` in https://github.com/apache/poi), licensed under the Apache License 2.0. They cover features the app must open or carry through untouched: charts, pivots, conditional formatting, data validation, macros, tables, shared formulas, strict OOXML, many merges and a hostile XML-bomb file.

`tests/unit/xlsx-roundtrip.test.ts` round-trips every file here except the hostile ones in `BAD_FIXTURES` (`tests/unit/helpers.ts`); the E2E specs open specific ones by name.
