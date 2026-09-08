# DOORS Monthly Production Freeze — 2026-09-08

This directory records the monthly-report Production hardening state installed manually in Apps Script on 2026-09-08 JST.

## Production targets
- Slides: 17qy7JYXMWmjhUbPrmvpURKOkT1TqsF8-jkM74jxw5-4
- Spreadsheet: 1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo
- Apps Script project: 1LQ4Mwhm52Kbowt8s7fJq15BSNBvYVBkik0faowMw-wUY4H54W5fWqcwj
- Production slide contract: 36 slides

## Installed files
- `DOORS_MONTHLY_PRODUCTION_FINALIZER_V16_1.gs`
  - additive finalizer
  - P.7 (`p8`) title/subtitle/table/insight from `GSCページ別データ`
  - P.30 (`SLIDES_API1535510610_0`) month labels and vertical-flip repair only when scaleY < 0
  - no acquisition, trigger, LLMO, Snapshot, or run_id mutation
- `DOORS_MONTHLY_SCHEDULER_HOTFIX_V14_REVIEWED.gs`
  - Production monthly entrypoint: `doorsScheduledMonthlyRunV14Reviewed`
  - resolves previous complete month in Asia/Tokyo
  - run order: `runDoorsMonthlyReportV21` -> `syncDoorsMonthlyInsightsV12` -> `syncDoorsKeywordProposalsToSlidesV1` -> `finalizeDoorsMonthlyProductionV16`

## Execution contract
Normal monthly Production execution uses only:

`doorsScheduledMonthlyRunV14Reviewed`

`diagnoseDoorsMonthlyProductionFinalizerV16` is diagnostic-only and is not part of the normal monthly operator procedure.

## Frozen assumptions / invariants
- STAGING deck is not a restoration source.
- The existing 36-slide Production deck is the Golden Master layout baseline.
- Spreadsheet COMMITTED SSOT is authoritative.
- Existing 2026-08 Production content is preserved; this freeze concerns future monthly output behavior.
- Existing LP/LLMO/SEO writer files are not replaced by this freeze.
- P.30 transform repair is idempotent: no change when vertical scale is already positive.

## Validation performed before freeze
- Production slide count/order checked at 36 slides.
- P.7 object IDs checked: `p8_i2`, `p8_i3`, `p8_i8`, `p8_i9`.
- P.7 table contract checked: 11 rows x 7 columns.
- 2026-08 GSC page SSOT replay reproduced: 9 declining pages; top-3 deltas -470/-300/-247; concentration 70.3% (~7割).
- P.30 body was visually repaired in Production and finalizer only repairs when negative scaleY is detected.
- Scheduler change is additive after existing monthly/insight/keyword steps.

## Rollback
Use the parent commit of this freeze commit to restore the repository state before v16.1, or restore only these two files from the desired earlier revision.

