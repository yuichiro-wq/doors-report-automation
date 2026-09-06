# DOORS Slide Headline Policy v1

## Rule

Slide headlines must communicate what the page shows, what changed, or the business/SEO insight.

Do not place operational or data-availability constraints in the headline, including:

- 未出力 / 未確定 / 未登録
- 直近確定月
- validation gate / 検証ゲート
- SSOT state
- 出力待ち / 再確定
- 推測・補完の都合
- 別ページを参照するという運用説明

These constraints belong in the subtitle, measurement note, or annotation.

## Production enforcement

`DOORS_GM_BRIDGE_V21.gs` runs `doorsApplyHeadlinePolicyV1_(reportMonth)` after the legacy monthly renderer and the article-performance renderer. This ordering prevents legacy rendering from reintroducing operational wording into the final headlines.

Currently enforced slide objectIds:

- `p8` — 流入LP
- `p11_channel_202607` — チャネル別流入
- `p13` — 読了
- `p23` — SEO最新ニュース

The policy is intentionally narrow and idempotent. It does not reorder slides, add/delete pages, or change data tables.
