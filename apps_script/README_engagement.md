# GA4のエンゲージメント（直帰率など）の取得：差し替え手順

DSPダッシュボード用の Apps Script（GA4＋GSC 自動取得）に、**エンゲージメント／直帰率の取得を追加した完成版**が `dsp_dashboard_auto.gs` です（2026-11-25）。
元のスクリプトの内容（`run_buildJson`・`doGet`・月次定例レポート用・訪問企業データ用など）は**そのまま含まれています**。末尾の「▼▼▼ エンゲージメントの追加」と、`fetchGa4_` の中の2か所だけが追加です。既存の出力項目は変えません。

> **GSCについて**：エンゲージメント・直帰率は GA4 の指標で、Search Console（GSC）にはありません。GA4 から取得します。

## 手順（ファイルを丸ごと差し替える）
1. Apps Script のプロジェクトで、元のスクリプトのファイル（`コード.gs` など、`run_buildJson` が入っているファイル）を開く
2. 中身を**すべて選択して削除**し、`apps_script/dsp_dashboard_auto.gs` の中身を**全部貼り付ける**（ファイルは1つのまま。ファイルを追加したり、別名のファイルに貼ったりしない）
3. 保存する。`CONFIG` の値（GA4のプロパティID・GSCのサイト・`driveFolderId` など）は、元のスクリプトと同じ値が入っています。元から変えている場合は、貼り付けたあとに合わせてください
4. `debug_engagement` を実行 → ログで SCH・IS の月次の直帰率・エンゲージメント時間・DSP流入の値を確認（初回は認証の許可が出る場合あり）
5. `run_buildJson` を実行して、エラーが出ないことを確認する（毎朝6時のトリガーは、そのまま動きます。「関数が削除されました」と出ていたトリガーは、関数が戻れば動きます。動かない場合は `setupDailyTrigger` を一度実行して作り直す）
6. 単独の確認用ファイルが欲しい場合は `run_buildEngagementJson` を実行（ドライブに `dsp_dashboard_engagement.json` が出る）

> **注意**：追加用のファイルだけを貼り付けて元のファイルを置き換えると、`run_buildJson` が無くなり「削除されました」のエラーで止まります（2026-11-25 に発生）。必ず**完成版を丸ごと**貼り付けてください。

## 目的
DSP広告（SCH・IS）の判定項目のうち、いまデータに無かった次の2つを取得します。
- **LPの直帰率の低下** → `bounceRate`（エンゲージしなかったセッションの割合）
- **LPの滞在時間の延長** → `avgEngagementSec`（1セッションあたりのエンゲージメント時間）／`engagementRate`・`viewsPerSession`（回遊）も取る

## 出力の形（`ga4[サイト].engagement`）
| 項目 | 内容 |
|---|---|
| `monthly` | サイト全体の月次 |
| `channelMonthly` | チャネル別（Display・Organic Search など）の月次 |
| `dspMonthly` | DSP流入（参照元／メディアに `admatrix` を含む）の月次。**SCH・ISのみ** |
| `dspLanding` | DSP流入のLP（ランディングページ）別の月次。**SCH・ISのみ** |

各行の指標：`sessions`／`engagedSessions`／`engagementRate`・`bounceRate`（0〜1）／`engagementSec`（合計）／`avgEngagementSec`（1セッションあたり）／`viewsPerSession`（1セッションあたりの閲覧ページ数）

## 設定（`ENGAGEMENT_CONFIG`）
- `enabled`：`false` にすると、エンゲージメントの取得を止められます（実行時間が足りなくなる場合）
- `dspSites`：DSP広告を実施しているサイト（既定：school／solution）
- `dspSourceMediumRegex`：DSP流入の判定（既定：`admatrix`。GA4の「セッションの参照元 / メディア」に含まれる文字列）

## 注意
- GA4の「直帰率」は、従来のUniversal Analyticsとは定義が違います（1ページだけ見て離脱した割合ではなく、**エンゲージしなかったセッションの割合**）。**エンゲージ＝10秒以上の滞在、またはコンバージョン、または2ページ以上の閲覧**。
- 月次レポートの「次ページ遷移率（約15%）」には、GA4のAPIに直接の指標がありません。当面は `viewsPerSession` と `bounceRate` で代替します。「2ページ以上見たセッションの割合」が必要なら、GA4の探索レポートか BigQuery エクスポートが必要です。
- DSPダッシュボード（HTML）は、このデータをまだ読みません（ソースがgit外のため改修の対象外）。出力JSONに項目が増えるだけなので、既存の表示は変わりません。ダッシュボードで使う場合は、別途ご相談ください。
- 取得期間は既存と同じ `CONFIG.lookbackDays`（800日）。Apps Scriptの実行時間（6分）が足りなくなる場合は、`ENGAGEMENT_CONFIG.enabled=false` で止められます。
- 動作確認：実際のGA4には接続できないため、模擬の応答で、構文・関数の存在・`fetchGa4_` の出力に `engagement` が入ることを確認しました。実データでの確認は `debug_engagement` で行ってください。
