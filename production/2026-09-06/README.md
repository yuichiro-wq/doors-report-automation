# DOORS Production Snapshot — 2026-09-06

This directory fixes the Production state after the 2026-08 monthly runtime verification completed on 2026-09-06.

## Runtime result

- `runDoorsMonthlyReportV21('2026-08')` completed successfully at 2026-09-06 09:40:45 JST.
- Production Deck remained 36 slides.
- Article LLMO 6 slides remained present.
- P.17–22 article-performance pages were updated.
- Operation panel was then updated and the bridge was extended so future runs write RUNNING / SUCCESS / BLOCK status automatically.

## Apps Script files fixed here

### DOORS_GM_BRIDGE_V21.gs
Exact Production bridge including operation-panel synchronization.

- bytes: 6532
- SHA-256: `a091ef6883d8c819367c7d23613f6ef242649aaa0aeabf83ce634e30787a5fb6`
- Git blob SHA: `7340e8e4e7b06b8c4a584b7f2e51d0cf9f48e145`

### DOORS_SEO_ARTICLE_PERFORMANCE_RENDERER_V11.gs
Exact renderer deployed for P.17–22.

- bytes: 14424
- SHA-256: `4c833a0e5742a9d43f0c4847022464e5f55fc92a861fb35026696e92fd35ca93`
- Git blob SHA: `82f4b6fcf98ffa9100dae3d7207cb3681e0cee98`

### DOORS_GA4_V21.gs
The exact 29,060-byte Production source is stored losslessly as two consecutive parts because of connector write-size handling:

1. `DOORS_GA4_V21.part1.gs`
2. `DOORS_GA4_V21.part2.gs`

Restore with:

```bash
cat DOORS_GA4_V21.part1.gs DOORS_GA4_V21.part2.gs > DOORS_GA4_V21.gs
```

Combined source:

- bytes: 29060
- SHA-256: `27096ce863208ad47bdc41965d5a32f51c08be911575d32323898531ebaa8dce`

Part Git blob SHAs:

- part1: `5999ef60ea44b65bfcfd04806b542054486a60f8`
- part2: `293999d0bad2eb08f8f4e1a9194bbedb7434d2f3`

## Production resources

- Apps Script project ID: `1LQ4Mwhm52Kbowt8s7fJq15BSNBvYVBkik0faowMw-wUY4H54W5fWqcwj`
- Spreadsheet ID: `1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo`
- Production Slides ID: `17qy7JYXMWmjhUbPrmvpURKOkT1TqsF8-jkM74jxw5-4`

## Important operational rule

Do not replace scheduler v1.4, LLMO v1.5, GSC configuration, SSOT IDs, or direct-run controls from this snapshot. This snapshot records only the three files changed/confirmed in the 2026-09-06 Production release sequence.
