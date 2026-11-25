# GA4のエンゲージメント（直帰率など）の取得：追加手順

DSPダッシュボード用の Apps Script（GA4＋GSC 自動取得）に、**エンゲージメント／直帰率の取得を追加**します（2026-11-25）。
追加分は別ファイル `engagement.gs` にまとめてあり、既存のスクリプトは **2か所の書き足しだけ**です。既存の出力項目は変えません。

> **GSCについて**：エンゲージメント・直帰率は GA4 の指標で、Search Console（GSC）にはありません。GA4 から取得します。

## 目的
DSP広告（SCH・IS）の判定項目のうち、いまデータに無かった次の2つを取得します。
- **LPの直帰率の低下** → `bounceRate`（エンゲージしなかったセッションの割合）
- **LPの滞在時間の延長** → `avgEngagementSec`（1セッションあたりのエンゲージメント時間）／`engagementRate`・`viewsPerSession`（回遊）も取る

## 手順
1. Apps Script のプロジェクトで、スクリプトファイルを追加（＋ → スクリプト）し、名前を `engagement` にして、`apps_script/engagement.gs` の中身を貼り付ける
2. 既存のメインのスクリプトの `fetchGa4_` を、次の2か所だけ書き足す

   **(a) `// 【追加 2026-10 訪問企業データ…】` のブロックの直後に追加：**
   ```js
   // 【追加 2026-11】エンゲージメント（直帰率・エンゲージメント率など）
   var engagement = null;
   try { engagement = fetchEngagement_(propertyId, start, end); }
   catch(err){ Logger.log('GA4 engagement失敗 %s: %s', propertyId, err); engagement = { error: String(err) }; }
   ```

   **(b) 関数の最後の `return { period: ... docodoco: docodoco };` に `engagement: engagement` を足す：**
   ```js
   return { period: start+'〜'+end, totals: totals, daily: daily, channels: channels,
            sources: sources, aiMonthly: aiMonthly, seoMonthly: seoMonthly, report: report,
            docodoco: docodoco, engagement: engagement };
   ```
3. `debug_engagement` を実行 → ログで SCH・IS の月次の直帰率・エンゲージメント時間・DSP流入の値を確認（初回は認証の許可が出る場合あり）
4. `run_buildEngagementJson` を実行すると、確認用ファイル `dsp_dashboard_engagement.json` がドライブに出る。以後は既存の `run_buildJson`（毎朝6時のトリガー）が、`dsp_dashboard_auto.json` の `ga4[サイト].engagement` に同じ内容を出力する

## 出力の形（`ga4[サイト].engagement`）
| 項目 | 内容 |
|---|---|
| `monthly` | サイト全体の月次 |
| `channelMonthly` | チャネル別（Display・Organic Search など）の月次 |
| `dspMonthly` | DSP流入（参照元／メディアに `admatrix` を含む）の月次。**SCH・ISのみ** |
| `dspLanding` | DSP流入のLP（ランディングページ）別の月次。**SCH・ISのみ** |

各行の指標：`sessions`／`engagedSessions`／`engagementRate`・`bounceRate`（0〜1）／`engagementSec`（合計）／`avgEngagementSec`（1セッションあたり）／`viewsPerSession`（1セッションあたりの閲覧ページ数）

## 設定（`engagement.gs` の `ENGAGEMENT_CONFIG`）
- `dspSites`：DSP広告を実施しているサイト（既定：school／solution）
- `dspSourceMediumRegex`：DSP流入の判定（既定：`admatrix`。GA4の「セッションの参照元 / メディア」に含まれる文字列）

## 注意
- GA4の「直帰率」は、従来のUniversal Analyticsとは定義が違います（1ページだけ見て離脱した割合ではなく、**エンゲージしなかったセッションの割合**）。**エンゲージ＝10秒以上の滞在、またはコンバージョン、または2ページ以上の閲覧**。
- 月次レポートの「次ページ遷移率（約15%）」には、GA4のAPIに直接の指標がありません。当面は `viewsPerSession` と `bounceRate` で代替します。「2ページ以上見たセッションの割合」が必要なら、GA4の探索レポートか BigQuery エクスポートが必要です。
- DSPダッシュボード（HTML）は、このデータをまだ読みません（ソースがgit外のため改修の対象外）。出力JSONに項目が増えるだけなので、既存の表示は変わりません。ダッシュボードで使う場合は、別途ご相談ください。
- 取得期間は既存と同じ `CONFIG.lookbackDays`（800日）。実行時間（Apps Scriptの6分制限）が足りなくなる場合は、`ENGAGEMENT_CONFIG.enabled=false` で止められます。
