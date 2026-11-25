# KPI・PDCA管理アプリ｜マーケティング本部

KPIツリーとPDCAを一体で管理するシングルHTMLアプリです。ブラウザで `KPI-PDCA管理アプリ.html` を開くだけで動作します（サーバー不要）。データはHTML内の `<script id="appData">` にJSONとして保持され、画面から編集・JSONエクスポート/インポートできます。

## データ置き場

- 入力データ（取込用Excel）：`C:\Users\612316\Box\事業推進Div□\DX推進課\生成AI\00_data`
  - 共通管理表（ROI_dashboard と共用）：`施策管理表.xlsx`・`マーケティング本部｜予実管理.xlsx`・`マーケティング本部｜コンテンツ制作｜本数管理.xlsx`・`マーケティング本部｜KPIツリー・施策評価.xlsx`
- 実績（`マーケティング本部｜{UN,SCH,MED,IS}｜実績.xlsm`）：**社内ドキュメント**（`C:\Users\612316\Box\マーケティンググループ\デジタルマーケティングDiv\マーケティング本部\社内ドキュメント`）から取り込む。00_data にある同名ファイルは古いコピーなので使わない（2026-10 確定）
- アプリの取込モーダルで上記フォルダのファイルを選ぶ（2026-10 に `source-data/` から移動）

## 構成

| ファイル / フォルダ | 内容 |
|---|---|
| `KPI-PDCA管理アプリ.html` | アプリ本体（UI・ロジック・データを1ファイルに内包） |
| `data/` | `kpi_pdca_data.json`（保存(JSON)の既定の保存先。git管理外）／**`candidates/`＝AI分析の候補ファイル**（`candidates_<事業>_<日付>.json` と、まとめた索引 `candidates.js`。アプリが開いたとき自動で読み込む。gitに置く） |
| `ai_analysis/` | AI分析：集計表の作成（`build_ai_input.py`）、候補ファイルの作成（`make_candidates*.py`・`candidates_lib.py`）、Word資料の作成（`make_*_docx.py`）、`input/`（AIに渡す集計表。個人情報なし） |
| `apps_script/` | DSPダッシュボード用 GA4＋GSC 取得スクリプトへの追加分（エンゲージメント・直帰率）と手順 |
| `docs/` | 仕様書・設計・引継ぎ資料・確認資料 |
| `00_data`（リポジトリ外） | データの元となるExcel（施策管理表・KPIツリーなど。実績.xlsm は社内ドキュメント）。上の「データ置き場」参照 |

### フォルダの構成（git・ローカル・Box で同じ）

このリポジトリ（GitHub）、ローカルの複製（`C:\Users\612316\Documents\GitHub\KPI-PDCA`）、Box の運用フォルダ（`生成AI\KPI-PDCA`）は、**同じ構成・同じ中身**にそろえます。違いが出るのは、gitに置かない次のものだけです。
- `data/kpi_pdca_data.json`・`backup/`（利用者のデータ。保存(JSON)が作る）
- 入力データ（`00_data`／`10_data`／社内ドキュメントの実績.xlsm など。個人情報を含む）
- `.claude/`（個人の設定）

更新の手順：ローカルで開発 → コミット → main へマージ → push → **Box の運用フォルダで `git pull`**（Boxのフォルダも同じリポジトリの複製）。Box のファイルを直接編集しない。

## 画面（タブ）

- 📊 全体サマリー … ファネルKPIの達成率・CV流入元・月次推移
- 🎯 施策別KPI … カテゴリ別の実施本数目標／実績・成果CV寄与
- 🌳 KPIツリー … チャネル別／ページ別セッション目標と実績
- 🔁 PDCAボード … Plan→Do→Check→Act の改善サイクル管理
- 🗂️ データ管理 … 各種実績・目標の入力／JSON入出力

## データモデル（appData 主要キー）

- `targets` / `actuals` … 事業別ファネルKPI（session/mql/sal/sql/受注）の目標・実績（＝全体目標）
- `kpiTree` … チャネル別・ページ別のセッション目標（＝個別KPI）
- `shisakuKpi` … カテゴリ別の実施本数目標（年間／月次）
- `seikaCv` … カテゴリ別の成果CV寄与
- `shisaku` … 施策管理表（PDCAの元となる施策一覧）
- `pdca` … PDCA改善サイクルの記録

## 開発メモ

- 今後の開発はこのgitリポジトリ上で行う。
- 対象事業コード：UN（企業ユニフォーム）/ SCH / MED / IS / SC。
- 会計年度：2026年度（2026-03〜2027-02）。
