# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository state

This is not a code project. There is no source code, build system, linter or test suite, and it is not a git repository. It holds sample datasets (`sim-data/data/*.xlsx`) and the rule tables (`sim-data/templates/*.md`) used to judge them. Any tooling would be created as new work.

Python on this machine is the Windows Store stub (`python` exits with code 49), so don't assume a working Python. Check for a real interpreter first, or use PowerShell and other available tools to read the `.xlsx` files.

## Layout

- `sim-data/data/` holds two parallel datasets:
  - `wafer_test_W07.xlsx`, `wafer_test_W21.xlsx`: semiconductor wafer test results.
  - `apple_A07.xlsx`, `apple_A25.xlsx`: apple (fruit) sorting data.
- `sim-data/templates/` holds the criteria tables, one per domain:
  - `합격기준표.md` is for wafers. It sets pass/fail.
  - `선별기준표.md` is for apples. It sets the grade.

## Data schema

Each workbook has a single sheet named after the file, with a header row and one row per item.

- **Wafer files** (616 die rows each, same die grid in W07 and W21, lot `LOT2611-07`). The columns are `lot_id, wafer_id, die_x, die_y, center_x_mm, center_y_mm, radius_mm, zone, vth_v, idsat_ma, ioff_na, rc_ohm, ox_thk_nm, crack, particle_cnt, bin, fail_type, pass_fail`. The judged output columns are `fail_type` (reason, `none` when passing) and `pass_fail` (`PASS`/`FAIL`).
- **Apple files** (600 rows each). The columns are `날짜, 사과번호, 가로칸, 세로칸, 중심거리, 위치, 품종, 무게, 지름, 당도, 멍, 벌레, 불량사유, 등급`. The judged output columns are `불량사유` (`없음` when passing) and `등급` (`상품`/`불량`).
- Strings are stored inline in the sheet XML. There is no `sharedStrings.xml`, which matters if you parse the files by hand.

Reading the files without Python: an `.xlsx` is a zip, so open it with `System.IO.Compression.ZipFile` in PowerShell and parse `xl/worksheets/sheet1.xml`.

## How the criteria tables work

Both tables use the same first-match-wins pattern, and any code or analysis built on them must keep it:

- Rows are checked in the order of the `순서` column. The first row whose condition fails sets the single reason for the judgment. If no row fails, the item passes (wafer reason `none`, apple reason `없음`).
- Wafer rows 6 are a tie. `idsat_ma` and `ox_thk_nm` both map to the `parametric` reason and have no relative order.
- The wafer table also gives each reason a machine name (`crack`, `open_short`, `leakage`, `vth_shift`, `particle`, `parametric`) and a display name. Keep both when producing output.
- Each table has an "집계 범위" (aggregation range) section. It lists the columns used to limit what gets aggregated, such as `radius_mm` for wafers and `중심거리` and `지름` for apples. The starting value is the data maximum, which means everything is included.
- Each table ends with a verification rule. Re-judging `wafer_test_W07.xlsx` must reproduce its `pass_fail` column exactly, and re-judging `apple_A07.xlsx` must reproduce its `등급` column exactly. Run that check after any change to the rules or the judging logic. `W21` and `A25` have the same output columns, but the tables state no check for them.

The templates are written in Korean and the column names in the data files are a mix of English (wafer) and Korean (apple). Match the exact column names from the tables.
