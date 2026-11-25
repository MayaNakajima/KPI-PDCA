/**
 * DSPダッシュボード用 GA4 取得スクリプトへの追加：エンゲージメント（直帰率など）の取得
 * ------------------------------------------------------------------
 * 目的：DSP広告の判定項目「LPの直帰率の低下」「LPの滞在時間の延長」に使うデータを、GA4から取る。
 *       これまでの出力には直帰率が無かった。
 *
 * ※「エンゲージメント」「直帰率」は GA4 の指標で、Search Console（GSC）にはありません。そのため GA4 から取得します。
 *    （GA4の直帰率＝エンゲージしなかったセッションの割合＝1−エンゲージメント率。
 *      「エンゲージしたセッション」＝10秒以上の滞在、またはコンバージョン、または2ページ以上の閲覧）
 *
 * 使い方（Apps Script のプロジェクトに、このファイルを「エンゲージメント」という名前のスクリプトファイルとして追加）
 *   1) 既存の fetchGa4_ の中を、README_engagement.md のとおり2か所だけ書き足す
 *      → 既存の出力（daily/channels/sources/aiMonthly/seoMonthly/report/docodoco/gsc）は変更しません。項目 engagement が増えるだけです
 *   2) 先に debug_engagement を実行して、ログで値を確認する（認証の確認も兼ねる）
 *   3) run_buildEngagementJson を実行すると、単独の確認用ファイル dsp_dashboard_engagement.json がドライブに出る
 *      （既存の run_buildJson を実行すれば、同じ内容が dsp_dashboard_auto.json の ga4[サイト].engagement に入る）
 *
 * 出力：ga4[サイト].engagement = {
 *    period: 取得期間,
 *    monthly:        [{ym:'2026/10', sessions, engagedSessions, engagementRate, bounceRate, engagementSec, avgEngagementSec, viewsPerSession}]   … サイト全体の月次
 *    channelMonthly: [{ym, ch:'Display', ...同じ指標}]                                                                                        … チャネル別の月次
 *    dspMonthly:     [{ym, sourceMedium:'admatrix / display', ...}]                                                                           … DSP流入の月次（DSP対象サイトのみ）
 *    dspLanding:     [{ym, page:'/', sourceMedium, ...}]                                                                                      … DSP流入のLP別の月次（DSP対象サイトのみ）
 *  }
 *  指標：sessions＝セッション数／engagedSessions＝エンゲージしたセッション数／engagementRate・bounceRate＝0〜1の割合
 *        engagementSec＝エンゲージメント時間の合計（秒）／avgEngagementSec＝1セッションあたり（秒）／viewsPerSession＝1セッションあたりの閲覧ページ数
 *  メモ：「次ページ遷移率」（月次レポートの約15%）は、GA4のAPIに直接の指標がありません。当面は viewsPerSession（回遊）と
 *        直帰率で代替します（セッション単位で「2ページ以上見た割合」が必要なら、GA4の探索レポートか BigQuery エクスポートが必要）。
 */

var ENGAGEMENT_CONFIG = {
  enabled: true,
  // DSP広告を実施しているサイト（CONFIG.ga4Properties のキー）。SCH と IS のみ
  dspSites: ['school.onward-cd.co.jp', 'solution.onward-cd.co.jp'],
  // DSP流入の判定（GA4の「セッションの参照元 / メディア」に含まれる文字列。例：admatrix / display）
  dspSourceMediumRegex: 'admatrix',
  rowLimit: 10000,        // 1リクエストの行数
  maxRows: 50000          // 1つの集計あたりの最大行数（JSONの肥大化の防止）
};

var ENG_METRICS = [
  { name: 'sessions' }, { name: 'engagedSessions' }, { name: 'engagementRate' },
  { name: 'bounceRate' }, { name: 'userEngagementDuration' }, { name: 'screenPageViewsPerSession' }
];

/** metricValues（ENG_METRICS の順）→ 出力の形 */
function engMetrics_(mv) {
  var sessions = +mv[0].value, dur = +mv[4].value;
  return {
    sessions: sessions,
    engagedSessions: +mv[1].value,
    engagementRate: +mv[2].value,
    bounceRate: +mv[3].value,
    engagementSec: dur,
    avgEngagementSec: sessions ? dur / sessions : 0,
    viewsPerSession: +mv[5].value
  };
}

/** yyyymm → 'yyyy/MM' */
function engYm_(v) { return v.slice(0, 4) + '/' + v.slice(4, 6); }

/** runReport をページ送りで全件取得する */
function engRun_(prop, req) {
  var rows = [], offset = 0, total = 0;
  do {
    req.limit = ENGAGEMENT_CONFIG.rowLimit;
    req.offset = offset;
    var r = AnalyticsData.Properties.runReport(req, prop);
    total = r.rowCount || 0;
    rows = rows.concat(r.rows || []);
    offset += ENGAGEMENT_CONFIG.rowLimit;
  } while (offset < total && offset < ENGAGEMENT_CONFIG.maxRows);
  return rows;
}

/** このプロパティがDSP対象サイトか */
function engIsDspProperty_(propertyId) {
  return ENGAGEMENT_CONFIG.dspSites.some(function (site) {
    return String(CONFIG.ga4Properties[site]) === String(propertyId);
  });
}

/** メイン：fetchGa4_ から呼ぶ（start/end は fetchGa4_ と同じ期間） */
function fetchEngagement_(propertyId, start, end) {
  if (!ENGAGEMENT_CONFIG.enabled) return null;
  var prop = 'properties/' + propertyId;
  var range = [{ startDate: start, endDate: end }];
  var out = { period: start + '〜' + end, monthly: [], channelMonthly: [], dspMonthly: [], dspLanding: [] };

  // (1) サイト全体の月次
  out.monthly = engRun_(prop, {
    dateRanges: range, dimensions: [{ name: 'yearMonth' }], metrics: ENG_METRICS,
    orderBys: [{ dimension: { dimensionName: 'yearMonth' } }]
  }).map(function (row) {
    var o = engMetrics_(row.metricValues); o.ym = engYm_(row.dimensionValues[0].value); return o;
  });

  // (2) チャネル別の月次
  out.channelMonthly = engRun_(prop, {
    dateRanges: range, dimensions: [{ name: 'yearMonth' }, { name: 'sessionDefaultChannelGroup' }], metrics: ENG_METRICS
  }).map(function (row) {
    var o = engMetrics_(row.metricValues);
    o.ym = engYm_(row.dimensionValues[0].value); o.ch = row.dimensionValues[1].value; return o;
  });

  // (3)(4) DSP流入（DSP対象サイトのみ）
  if (engIsDspProperty_(propertyId)) {
    var dspFilter = { filter: { fieldName: 'sessionSourceMedium',
      stringFilter: { matchType: 'PARTIAL_REGEXP', value: ENGAGEMENT_CONFIG.dspSourceMediumRegex, caseSensitive: false } } };
    try {
      out.dspMonthly = engRun_(prop, {
        dateRanges: range, dimensions: [{ name: 'yearMonth' }, { name: 'sessionSourceMedium' }], metrics: ENG_METRICS,
        dimensionFilter: dspFilter
      }).map(function (row) {
        var o = engMetrics_(row.metricValues);
        o.ym = engYm_(row.dimensionValues[0].value); o.sourceMedium = row.dimensionValues[1].value; return o;
      });
    } catch (err) { Logger.log('GA4 dspMonthly失敗 %s: %s', propertyId, err); }
    try {
      out.dspLanding = engRun_(prop, {
        dateRanges: range,
        dimensions: [{ name: 'yearMonth' }, { name: 'landingPage' }, { name: 'sessionSourceMedium' }],
        metrics: ENG_METRICS, dimensionFilter: dspFilter
      }).map(function (row) {
        var o = engMetrics_(row.metricValues);
        o.ym = engYm_(row.dimensionValues[0].value); o.page = row.dimensionValues[1].value;
        o.sourceMedium = row.dimensionValues[2].value; return o;
      });
    } catch (err) { Logger.log('GA4 dspLanding失敗 %s: %s', propertyId, err); }
  }
  return out;
}

/** 確認用：DSP対象サイトの直近月の値をログに出す（先にこれを実行して認証・値を確認する） */
function debug_engagement() {
  var start = isoDaysAgo(120), end = isoToday();
  ENGAGEMENT_CONFIG.dspSites.forEach(function (site) {
    var pid = CONFIG.ga4Properties[site];
    var e = fetchEngagement_(pid, start, end);
    Logger.log('== %s（property %s）期間 %s', site, pid, e.period);
    e.monthly.forEach(function (m) {
      Logger.log('  サイト全体 %s  セッション%s  エンゲージ率%s%%  直帰率%s%%  平均エンゲージ%s秒  閲覧/セッション%s',
        m.ym, m.sessions, (m.engagementRate * 100).toFixed(1), (m.bounceRate * 100).toFixed(1), m.avgEngagementSec.toFixed(1), m.viewsPerSession.toFixed(2));
    });
    e.dspMonthly.forEach(function (m) {
      Logger.log('  DSP %s  %s  セッション%s  エンゲージ率%s%%  直帰率%s%%  平均エンゲージ%s秒  閲覧/セッション%s',
        m.ym, m.sourceMedium, m.sessions, (m.engagementRate * 100).toFixed(1), (m.bounceRate * 100).toFixed(1), m.avgEngagementSec.toFixed(1), m.viewsPerSession.toFixed(2));
    });
    var top = e.dspLanding.slice().sort(function (a, b) { return b.sessions - a.sessions; }).slice(0, 5);
    top.forEach(function (m) {
      Logger.log('  DSPのLP上位 %s  %s  セッション%s  直帰率%s%%', m.ym, m.page, m.sessions, (m.bounceRate * 100).toFixed(1));
    });
  });
}

/** 単独で確認用ファイル dsp_dashboard_engagement.json をドライブに出力する（既存の run_buildJson とは独立） */
function run_buildEngagementJson() {
  var start = isoDaysAgo(CONFIG.lookbackDays), end = isoToday();
  var out = { type: 'engagement', generated: new Date().toISOString(), ga4: {} };
  for (var site in CONFIG.ga4Properties) {
    try { out.ga4[site] = fetchEngagement_(CONFIG.ga4Properties[site], start, end); }
    catch (err) { Logger.log('GA4 engagement失敗 %s: %s', site, err); }
  }
  var name = 'dsp_dashboard_engagement.json', text = JSON.stringify(out);
  var folder = CONFIG.driveFolderId ? DriveApp.getFolderById(CONFIG.driveFolderId) : DriveApp.getRootFolder();
  var it = folder.getFilesByName(name);
  if (it.hasNext()) { it.next().setContent(text); } else { folder.createFile(name, text, MimeType.PLAIN_TEXT); }
  Logger.log('生成完了: %s サイト → %s', Object.keys(out.ga4).length, name);
}
