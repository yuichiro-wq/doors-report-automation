# DOORS Production visual contract — P11 / P17-P22

Frozen after visual regression recovery on 2026-09-08.

## P11 channel page
- Purpose: show the full Japan /doors/ channel mix, not an AI-only interpretation.
- Compare 2026-07 vs 2026-08 using `ga4-channel-japan-doors-lpqs-v2` rows from `チャネル別月次`.
- Include Organic Search, Direct, Referral, Organic Social, AI Assistant, Email, Unassigned, Paid Search, Organic Video.
- AI Assistant is one GA4 channel. Do not substitute the separate AI / LLM source-contract KPI from P12.
- Required visual: channel comparison graph plus a concise bottom insight.
- Approved title pattern: `チャネル別流入｜7月・8月の流入構成をチャネル別に比較`.

## P17-P22 article performance
- Historical completed design is the visual baseline: `DOORS_月次レポート_2026年8月実績(1).pptx`.
- Use a full-width table with dark navy header, not sparse body text and not 2x3 cards.
- Preserve current 36-slide architecture and object IDs:
  - p20, p20_detail_2, p20_detail_3
  - p21, p21_detail_2, p21_detail_3
- Current page split remains 6 rows x 3 pages for 新規 and 6 rows x 3 pages for リライト.
- New article columns: 記事 / 8月PV / 7月PV / 増減 / 対策KW順位.
- Rewrite columns: 記事 / 8月PV / 7月PV / 増減 / GSC平均順位.
- Positive change is visually distinguished from negative change.
- Each page carries a short bottom insight strip (伸長 / 要確認 / 優先改善).

## Regression prohibition
- Do not run or deploy a renderer that removes the historical reference design and replaces P17-P22 with plain multiline text.
- Any monthly renderer must preserve this visual contract before Production promotion.
