/**
 * DSPダッシュボード用 GA4 + Search Console 自動取得スクリプト（Google Apps Script）
 * ------------------------------------------------------------------
 * 役割：GA4とサーチコンソール(GSC)のデータを定期取得し、ダッシュボードが読める
 *       JSON を生成します。出力は (1) ウェブアプリのURL（☁️最新を取得で読む）と
 *       (2) Googleドライブのファイル（📥 JSON読み込みで読む）の両方に対応。
 *
 * 【2026-09 追加】月次定例レポート用の集計（ga4[サイト].report）
 *   ・年月×チャネルのセッション数（P5 チャネル別セッション数推移）
 *   ・月別のコラム／事例ページのパフォーマンス（P8/P9 上位10位）
 *   ・月別のランディングページ×参照元/メディア（P6）
 *   追加部分はファイル末尾の「▼▼▼ 月次定例レポート用の追加」にまとめています。
 *
 * 【2026-10 追加】訪問企業データ（ga4[サイト].docodoco）
 *   ・どこどこJP（BtoB2設定）の組織名・業種大分類・売上高 ＋ uSonarの社名・法人番号・資本金
 *   ・企業×閲覧日×閲覧ページ（visits）と、業種×年月（industryMonthly）
 *   追加部分はファイル末尾の「▼▼▼ 訪問企業データの追加」にまとめています。
 *
 * 【2026-11 追加】エンゲージメント（ga4[サイト].engagement）
 *   ・直帰率・エンゲージメント率・エンゲージメント時間・閲覧ページ数（サイト全体／チャネル別の月次）
 *   ・DSP広告（SCH・IS）の流入の月次・LP別の月次 … DSPの判定項目「LPの直帰率の低下」「LPの滞在時間の延長」用
 *   ・デバイス別（PC＝desktop／SP＝mobile）の月次（サイト全体・DSP流入）… PCとSPの予算配分の見直しの効果を見る用
 *   追加部分はファイル末尾の「▼▼▼ エンゲージメントの追加」にまとめています。
 *   ※「エンゲージメント」「直帰率」は GA4 の指標で、Search Console（GSC）にはありません。
 *   既存の出力項目（daily/channels/sources/aiMonthly/seoMonthly/report/docodoco/gsc）は変更していません。
 *
 * 事前準備（詳細は手順書 AppsScript_setup_guide.md 参照）：
 *   1) GA4プロパティとGSCサイトに閲覧権限のあるGoogleアカウントで作成
 *   2) 拡張サービス「Analytics Data API」をON（サービス名: AnalyticsData）
 *   3) appsscript.json に GSC 用スコープを追加（手順書参照）
 *   4) 下の CONFIG を自社の値に書き換え
 *   5) run_buildJson を一度手動実行して認証 → トリガーで定期実行
 */

// ============ CONFIG（ここだけ書き換え） ============
var CONFIG = {
  // GA4：サイト識別子 → GA4プロパティID（数字のみ）。サイト識別子はドメインを推奨（ダッシュボードが事業名に自動変換）
  ga4Properties: {
    'uniform.onward-cd.co.jp':  '456801598',   // UN
    'school.onward-cd.co.jp':   '456543598',   // SCH
    'medical.onward-cd.co.jp':  '434136175',   // MED
    'solution.onward-cd.co.jp': '444043577',   // IS
    'onward-cd.co.jp':          '534387764'    // コーポレート
  },
  // GSC：登録プロパティのURL（末尾スラッシュ含む形 or sc-domain: 形式）
  gscSites: [
    'https://uniform.onward-cd.co.jp/',
    'https://school.onward-cd.co.jp/',
    'https://medical.onward-cd.co.jp/',
    'https://solution.onward-cd.co.jp/',
    'https://onward-cd.co.jp/'
  ],
  lookbackDays: 800,                // GA4の取得期間（日数）。配信前（2025/05〜）との比較のため長め
  gscLookbackDays: 480,             // GSCの取得期間（日数）。GSC自体の保存は約16か月（それ以前は返らない）
  queryMonthlyRowLimit: 500,        // 月別クエリ（指名検索の月次推移用）の取得件数／月
  ga4ConversionMetric: 'keyEvents', // 新しいGA4は 'keyEvents'。0件なら 'conversions' に変更
  brandRegex: /(オンワード|onward|raffiria|ラフィーリア|インサイトセールス|insight)/i, // 指名(ブランド)判定
  // 外部AIアシスタント（ChatGPT/Perplexity等）からの流入判定に使う参照元ホスト
  aiSourceRegex: /(chatgpt\.com|chat\.openai\.com|openai\.com|perplexity\.ai|gemini\.google\.com|bard\.google\.com|copilot\.microsoft\.com|bing\.com\/chat|you\.com|claude\.ai|poe\.com|phind\.com)/i,
  driveFileName: 'dsp_dashboard_auto.json', // ドライブ保存名（📥用）
  driveFolderId: ''                 // 空ならマイドライブ直下。特定フォルダに置くならフォルダIDを指定
};
// ===================================================

function isoDaysAgo(n){ var d=new Date(); d.setDate(d.getDate()-n); return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }
function isoToday(){ return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); }

/** メイン：JSONを組み立ててドライブへ保存（トリガーはこれを呼ぶ） */
function run_buildJson(){
  var payload = buildPayload();
  saveToDrive_(JSON.stringify(payload));
  Logger.log('生成完了: GA4 %s サイト / GSC %s サイト',
    Object.keys(payload.ga4).length, Object.keys(payload.gsc).length);
  return payload;
}

/** ウェブアプリ：デプロイ後のURLに GET すると JSON を返す（☁️最新を取得 用） */
function doGet(e){
  var payload = buildPayload();
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildPayload(){
  var out = { type:'auto', generated: new Date().toISOString(), ga4:{}, gsc:{} };
  var start = isoDaysAgo(CONFIG.lookbackDays), end = isoToday();
  var gscStart = isoDaysAgo(CONFIG.gscLookbackDays);
  // GA4
  for (var site in CONFIG.ga4Properties){
    try { out.ga4[site] = fetchGa4_(CONFIG.ga4Properties[site], start, end); }
    catch(err){ Logger.log('GA4失敗 %s: %s', site, err); }
  }
  // GSC
  for (var i=0;i<CONFIG.gscSites.length;i++){
    var url = CONFIG.gscSites[i];
    try {
      var key = url.replace(/^https?:\/\//,'').replace(/^sc-domain:/,'').replace(/\/$/,'');
      out.gsc[key] = fetchGsc_(url, gscStart, end);
    } catch(err){ Logger.log('GSC失敗 %s: %s', url, err); }
  }
  return out;
}

/** GA4：Analytics Data API（拡張サービス AnalyticsData） */
function fetchGa4_(propertyId, start, end){
  var cvMetric = CONFIG.ga4ConversionMetric;
  // 日次 sessions / users
  var rDaily = AnalyticsData.Properties.runReport({
    dateRanges:[{startDate:start, endDate:end}],
    dimensions:[{name:'date'}],
    metrics:[{name:'sessions'},{name:'totalUsers'}],
    orderBys:[{dimension:{dimensionName:'date'}}],
    limit: 1000
  }, 'properties/'+propertyId);
  var daily = (rDaily.rows||[]).map(function(row){
    var d = row.dimensionValues[0].value; // yyyymmdd
    return { date: d.slice(0,4)+'-'+d.slice(4,6)+'-'+d.slice(6,8),
             sessions:+row.metricValues[0].value, users:+row.metricValues[1].value };
  });
  // 合計 sessions / users / conversions
  var rTot = AnalyticsData.Properties.runReport({
    dateRanges:[{startDate:start, endDate:end}],
    metrics:[{name:'sessions'},{name:'totalUsers'},{name:cvMetric}]
  }, 'properties/'+propertyId);
  var tv = (rTot.rows&&rTot.rows[0]) ? rTot.rows[0].metricValues : [{value:0},{value:0},{value:0}];
  var totals = { sessions:+tv[0].value, users:+tv[1].value, conversions:+tv[2].value };
  // チャネル別 sessions / conversions
  var rCh = AnalyticsData.Properties.runReport({
    dateRanges:[{startDate:start, endDate:end}],
    dimensions:[{name:'sessionDefaultChannelGroup'}],
    metrics:[{name:'sessions'},{name:cvMetric}],
    orderBys:[{metric:{metricName:'sessions'}, desc:true}],
    limit: 20
  }, 'properties/'+propertyId);
  var channels = (rCh.rows||[]).map(function(row){
    return { name: row.dimensionValues[0].value,
             sessions:+row.metricValues[0].value, conversions:+row.metricValues[1].value };
  });
  // 参照元別 sessions（AI流入=外部AIアシスタントの判定に使用）
  var sources = [];
  try {
    var rSrc = AnalyticsData.Properties.runReport({
      dateRanges:[{startDate:start, endDate:end}],
      dimensions:[{name:'sessionSource'}],
      metrics:[{name:'sessions'},{name:cvMetric}],
      orderBys:[{metric:{metricName:'sessions'}, desc:true}],
      limit: 300
    }, 'properties/'+propertyId);
    sources = (rSrc.rows||[]).map(function(row){
      var nm = row.dimensionValues[0].value;
      return { name:nm, sessions:+row.metricValues[0].value, conversions:+row.metricValues[1].value,
               ai: CONFIG.aiSourceRegex.test(nm) };
    });
  } catch(err){ Logger.log('GA4 sources失敗 %s: %s', propertyId, err); }
  // 外部AI流入の月次（yearMonth × 参照元、AIホストのみ）
  var aiMonthly = [];
  try {
    var rAiM = AnalyticsData.Properties.runReport({
      dateRanges:[{startDate:start, endDate:end}],
      dimensions:[{name:'yearMonth'},{name:'sessionSource'}],
      metrics:[{name:'sessions'},{name:cvMetric}],
      dimensionFilter:{ filter:{ fieldName:'sessionSource',
        stringFilter:{ matchType:'PARTIAL_REGEXP',
          value:'chatgpt|openai|perplexity|gemini|bard|copilot|bing\\.com/chat|you\\.com|claude\\.ai|poe\\.com|phind' } } },
      limit: 2000
    }, 'properties/'+propertyId);
    var aiMap = {};
    (rAiM.rows||[]).forEach(function(row){
      var ym = row.dimensionValues[0].value; // yyyymm
      var key = ym.slice(0,4)+'/'+ym.slice(4,6);
      aiMap[key] = aiMap[key] || { ym:key, sessions:0, conversions:0 };
      aiMap[key].sessions += +row.metricValues[0].value;
      aiMap[key].conversions += +row.metricValues[1].value;
    });
    aiMonthly = Object.keys(aiMap).sort().map(function(k){ return aiMap[k]; });
  } catch(err){ Logger.log('GA4 aiMonthly失敗 %s: %s', propertyId, err); }
  // 検索エンジン流入(Organic Search)の月次
  var seoMonthly = [];
  try {
    var rSeoM = AnalyticsData.Properties.runReport({
      dateRanges:[{startDate:start, endDate:end}],
      dimensions:[{name:'yearMonth'},{name:'sessionDefaultChannelGroup'}],
      metrics:[{name:'sessions'}],
      limit: 2000
    }, 'properties/'+propertyId);
    var seoMap = {};
    (rSeoM.rows||[]).forEach(function(row){
      var ym = row.dimensionValues[0].value, ch = row.dimensionValues[1].value;
      if (ch !== 'Organic Search') return;
      var key = ym.slice(0,4)+'/'+ym.slice(4,6);
      seoMap[key] = (seoMap[key]||0) + (+row.metricValues[0].value);
    });
    seoMonthly = Object.keys(seoMap).sort().map(function(k){ return { ym:k, sessions:seoMap[k] }; });
  } catch(err){ Logger.log('GA4 seoMonthly失敗 %s: %s', propertyId, err); }
  // 【追加】月次定例レポート用
  var report = null;
  try { report = fetchMonthlyReport_(propertyId, start, end); }
  catch(err){ Logger.log('GA4 report失敗 %s: %s', propertyId, err); }
  // 【追加 2026-10】訪問企業データ（どこどこJP＋uSonar）
  var docodoco = null;
  try { docodoco = fetchDocodoco_(propertyId, end); }
  catch(err){ Logger.log('GA4 docodoco失敗 %s: %s', propertyId, err); docodoco = { error: String(err) }; }
  // 【追加 2026-11】エンゲージメント（直帰率・エンゲージメント率など）
  var engagement = null;
  try { engagement = fetchEngagement_(propertyId, start, end); }
  catch(err){ Logger.log('GA4 engagement失敗 %s: %s', propertyId, err); engagement = { error: String(err) }; }
  return { period: start+'〜'+end, totals: totals, daily: daily, channels: channels,
           sources: sources, aiMonthly: aiMonthly, seoMonthly: seoMonthly, report: report,
           docodoco: docodoco, engagement: engagement };
}

/** GSC：Search Console API（UrlFetch + OAuthトークン） */
function fetchGsc_(siteUrl, start, end){
  var base = 'https://www.googleapis.com/webmasters/v3/sites/'+encodeURIComponent(siteUrl)+'/searchAnalytics/query';
  function q(body){
    var res = UrlFetchApp.fetch(base, {
      method:'post', contentType:'application/json',
      headers:{ Authorization:'Bearer '+ScriptApp.getOAuthToken() },
      payload: JSON.stringify(body), muteHttpExceptions:true
    });
    var code = res.getResponseCode();
    if (code!==200) throw new Error('GSC '+code+': '+res.getContentText().slice(0,200));
    return JSON.parse(res.getContentText()).rows || [];
  }
  var daily = q({startDate:start, endDate:end, dimensions:['date'], rowLimit:1000}).map(function(r){
    return { date:r.keys[0], clk:r.clicks, imp:r.impressions, ctr:r.ctr, pos:r.position };
  });
  var queries = q({startDate:start, endDate:end, dimensions:['query'], rowLimit:1000}).map(function(r){
    return { q:r.keys[0], clk:r.clicks, imp:r.impressions, ctr:r.ctr, pos:r.position,
             brand: CONFIG.brandRegex.test(r.keys[0]) };
  });
  var pages = q({startDate:start, endDate:end, dimensions:['page'], rowLimit:100}).map(function(r){
    return { page:r.keys[0], clk:r.clicks, imp:r.impressions, ctr:r.ctr, pos:r.position };
  });
  var device = q({startDate:start, endDate:end, dimensions:['device'], rowLimit:10}).map(function(r){
    return { name:r.keys[0], clk:r.clicks, imp:r.impressions, ctr:r.ctr };
  });
  // 月別の上位クエリ（指名検索の月次推移用）。ダッシュボード側で月単位に蓄積するため、GSCの保存期間を過ぎても消えない
  var queryMonthly = [];
  try {
    monthList_(start, end).forEach(function(m){
      var rows = q({startDate:m.start, endDate:m.end, dimensions:['query'], rowLimit:CONFIG.queryMonthlyRowLimit});
      queryMonthly.push({ ym:m.ym, start:m.start, end:m.end, complete:m.complete,
        queries: rows.map(function(r){ return { q:r.keys[0], clk:r.clicks, imp:r.impressions }; }) });
    });
  } catch(err){ Logger.log('GSC queryMonthly失敗 %s: %s', siteUrl, err); }
  function agg(arr){ return { imp:arr.reduce(function(s,x){return s+x.imp;},0),
                              clk:arr.reduce(function(s,x){return s+x.clk;},0), n:arr.length }; }
  var brand = agg(queries.filter(function(x){return x.brand;}));
  var generic = agg(queries.filter(function(x){return !x.brand;}));
  var siteKey = siteUrl.replace(/^https?:\/\//,'').replace(/^sc-domain:/,'').replace(/\/$/,'');
  return { site:siteKey, period:start+'〜'+end, daily:daily,
           queries:queries.slice(0,300), pages:pages.slice(0,40),
           device:device, brand:brand, generic:generic, queryMonthly:queryMonthly };
}

/** start〜end を暦月に区切る（yyyy-MM-dd）。当月など途中の月は complete:false */
function monthList_(start, end){
  var out = [], s = new Date(start+'T00:00:00'), e = new Date(end+'T00:00:00');
  var cur = new Date(s.getFullYear(), s.getMonth(), 1);
  while (cur <= e){
    var mStart = cur < s ? s : cur;
    var last = new Date(cur.getFullYear(), cur.getMonth()+1, 0);
    var mEnd = last > e ? e : last;
    var f = function(d){ return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); };
    out.push({ ym: Utilities.formatDate(cur, 'Asia/Tokyo', 'yyyy/MM'), start:f(mStart), end:f(mEnd),
               complete: mStart.getDate()===1 && mEnd.getTime()===last.getTime() });
    cur = new Date(cur.getFullYear(), cur.getMonth()+1, 1);
  }
  return out;
}

/** ドライブ保存（📥 JSON読み込み 用の共有ファイル） */
function saveToDrive_(text){
  var name = CONFIG.driveFileName;
  var folder = CONFIG.driveFolderId ? DriveApp.getFolderById(CONFIG.driveFolderId) : DriveApp.getRootFolder();
  var it = folder.getFilesByName(name);
  if (it.hasNext()) { it.next().setContent(text); }
  else { folder.createFile(name, text, MimeType.PLAIN_TEXT); }
}

/** 定期実行トリガーを作成（例：毎日午前6時） */
function setupDailyTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction()==='run_buildJson') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('run_buildJson').timeBased().everyDays(1).atHour(6).create();
  Logger.log('毎日6時のトリガーを作成しました');
}

// ===================================================================
//  ▼▼▼ 月次定例レポート用の追加（2026-09 追加） ▼▼▼
//  出力：payload.ga4[サイト].report = {
//    channelMonthly: [{ym:'2026/08', ch:'Organic Search', sessions:123}, ...]   … P5 チャネル別セッション数推移
//    months: [{ym, start, end, complete,
//              pages:[{path,title,views,sessions,viewsPerSession,users,activeUsers,engagementSec,scrollUsers}],  … P8/P9 上位10位
//              landing:[{page,sourceMedium(=ユーザーの最初の参照元/メディア),sessions,activeUsers,newUsers,engagementSec}], landingTotal:{...}（当月の全体合計）}]   … P6・月間セッション数
//  }
// ===================================================================
var REPORT_CONFIG = {
  monthsBack: 3,                 // ページ別・ランディングページ別を取る月数（当月を含まない完了月の数）
  pagePathRegex: 'column|case',  // コラム・事例ページの判定（ページパスに含む文字）
  pageLimit: 300,                // 1か月あたりのページ行数の上限
  landingLimit: 30               // 1か月あたりのランディングページ×参照元の行数
};

function fetchMonthlyReport_(propertyId, start, end){
  var prop = 'properties/'+propertyId;
  var out = { channelMonthly: [], months: [] };
  // (1) 年月 × チャネル
  var rCh = AnalyticsData.Properties.runReport({
    dateRanges:[{startDate:start, endDate:end}],
    dimensions:[{name:'yearMonth'},{name:'sessionDefaultChannelGroup'}],
    metrics:[{name:'sessions'}],
    limit: 5000
  }, prop);
  out.channelMonthly = (rCh.rows||[]).map(function(row){
    var ym = row.dimensionValues[0].value;
    return { ym: ym.slice(0,4)+'/'+ym.slice(4,6), ch: row.dimensionValues[1].value, sessions: +row.metricValues[0].value };
  });
  // (2) 月ごとのページ別・ランディングページ別
  var now = new Date();
  var mStart = new Date(now.getFullYear(), now.getMonth()-REPORT_CONFIG.monthsBack, 1);
  var list = monthList_(Utilities.formatDate(mStart,'Asia/Tokyo','yyyy-MM-dd'), end);
  var pathFilter = { filter:{ fieldName:'unifiedPagePathScreen',
    stringFilter:{ matchType:'PARTIAL_REGEXP', value: REPORT_CONFIG.pagePathRegex } } };
  list.forEach(function(m){
    var mo = { ym:m.ym, start:m.start, end:m.end, complete:m.complete, pages:[], landing:[], landingTotal:null };
    try {
      var rP = AnalyticsData.Properties.runReport({
        dateRanges:[{startDate:m.start, endDate:m.end}],
        dimensions:[{name:'unifiedPagePathScreen'},{name:'pageTitle'}],
        metrics:[{name:'screenPageViews'},{name:'sessions'},{name:'screenPageViewsPerSession'},
                 {name:'totalUsers'},{name:'activeUsers'},{name:'userEngagementDuration'}],
        dimensionFilter: pathFilter,
        orderBys:[{metric:{metricName:'screenPageViews'}, desc:true}],
        limit: REPORT_CONFIG.pageLimit
      }, prop);
      // 90%スクロール（拡張計測の scroll イベント）のユーザー数
      var rS = AnalyticsData.Properties.runReport({
        dateRanges:[{startDate:m.start, endDate:m.end}],
        dimensions:[{name:'unifiedPagePathScreen'},{name:'pageTitle'}],
        metrics:[{name:'totalUsers'}],
        dimensionFilter:{ andGroup:{ expressions:[ pathFilter,
          { filter:{ fieldName:'eventName', stringFilter:{ matchType:'EXACT', value:'scroll' } } } ] } },
        limit: REPORT_CONFIG.pageLimit*2
      }, prop);
      var scroll = {};
      (rS.rows||[]).forEach(function(row){
        scroll[row.dimensionValues[0].value+'\t'+row.dimensionValues[1].value] = +row.metricValues[0].value;
      });
      mo.pages = (rP.rows||[]).map(function(row){
        var p = row.dimensionValues[0].value, t = row.dimensionValues[1].value, v = row.metricValues;
        return { path:p, title:t, views:+v[0].value, sessions:+v[1].value, viewsPerSession:+v[2].value,
                 users:+v[3].value, activeUsers:+v[4].value, engagementSec:+v[5].value,
                 scrollUsers: scroll[p+'\t'+t]||0 };
      });
    } catch(err){ Logger.log('GA4 pages失敗 %s %s: %s', propertyId, m.ym, err); }
    try {
      var rL = AnalyticsData.Properties.runReport({
        dateRanges:[{startDate:m.start, endDate:m.end}],
        dimensions:[{name:'landingPage'},{name:'firstUserSourceMedium'}],  // GA4探索画面と同じ「ランディング ページ」×「ユーザーの最初の参照元 / メディア」   // 広告パラメータ等を除いたページ単位（GA4画面の「ランディング ページ」と同じ）
        metrics:[{name:'sessions'},{name:'activeUsers'},{name:'newUsers'},{name:'userEngagementDuration'}],
        orderBys:[{metric:{metricName:'sessions'}, desc:true}],
        metricAggregations:['TOTAL'],
        limit: REPORT_CONFIG.landingLimit
      }, prop);
      mo.landing = (rL.rows||[]).map(function(row){
        var v = row.metricValues;
        return { page:row.dimensionValues[0].value, sourceMedium:row.dimensionValues[1].value,
                 sessions:+v[0].value, activeUsers:+v[1].value, newUsers:+v[2].value, engagementSec:+v[3].value };
      });
      var tt = rL.totals && rL.totals[0] ? rL.totals[0].metricValues : null;
      if (tt) mo.landingTotal = { sessions:+tt[0].value, activeUsers:+tt[1].value, newUsers:+tt[2].value, engagementSec:+tt[3].value };
    } catch(err){ Logger.log('GA4 landing失敗 %s %s: %s', propertyId, m.ym, err); }
    out.months.push(mo);
  });
  return out;
}
// ▲▲▲ 月次定例レポート用の追加 ここまで ▲▲▲

// ===================================================================
//  ▼▼▼ 訪問企業データの追加（2026-10 追加） ▼▼▼
//  前提：GTMタグ「イベント_GA4_どこどこ」（どこどこJP BtoB2設定）と uSonar ライブアクセスが、
//        組織名などをGA4のユーザープロパティ（範囲：ユーザー）として送信していること
//
//  出力：payload.ga4[サイト].docodoco = {
//    dims:      実際に使ったGA4ディメンション名（プロパティごとに自動判定した結果）
//    period:    取得期間
//    rowCount / truncated: GA4側の総行数／上限で切ったかどうか
//    visits: [{date, org, industry, sales, usCompany, houjin, capital, path, views, sessions}]  … 企業×閲覧日×閲覧ページ
//       org/industry/sales = どこどこJP（組織名・業種大分類・売上高）
//       usCompany/houjin/capital = uSonar（社名・法人番号・資本金）
//    industryMonthly: [{ym, industry, sessions}]   … どこどこJPの業種大分類×年月
//  }
//  確認用：debug_listCustomDims を実行すると、各プロパティのカスタムディメンション一覧がログに出ます
// ===================================================================
var DOCODOCO_CONFIG = {
  enabled: true,
  lookbackDays: 90,     // visits の取得期間（日）。長くするとJSONが大きくなります
  // 取得する項目。api は候補を上から順に試し、プロパティに登録されている最初のものを使用。
  // 見つからない場合は、説明欄が src に一致するカスタムディメンションの中から、表示名（ui）で探します。
  dims: {
    org:       { api:['customUser:組織名'],                     src:/どこどこ/,  ui:[/^組織名$/] },
    industry:  { api:['customUser:業種大分類'],                 src:/どこどこ/,  ui:[/^業種大分類$/] },
    // GTMは「売上高」で送信。GA4に「売上金」で登録されているプロパティもあるため両方を候補にする
    sales:     { api:['customUser:売上高','customUser:売上金'], src:/どこどこ/,  ui:[/^売上(高|金)$/] },
    usCompany: { api:['customUser:us_company_name'],            src:/uSonar/i,  ui:[/^社名$/] },
    houjin:    { api:['customUser:us_corporate_number'],        src:/uSonar/i,  ui:[/^法人番号$/] },
    capital:   { api:['customUser:us_capital'],                 src:/uSonar/i,  ui:[/^資本金/] }
  },
  // ISP回線などを除外したい場合の組織名の正規表現（例：/(NTT|KDDI|ソフトバンク).*回線/）。不要なら null
  excludeOrgRegex: /(オンワード樫山|オンワードコーポレートデザイン)/,
  // 除外する法人番号（社内・グループ会社のアクセスなど）。例：['7010001108009']
  excludeHoujin: ['7010001108009', '8010601042176'],   // （株）オンワード樫山、（株）オンワードコーポレートデザイン
  // 「値なし」として空欄に置き換える値（uSonarは未判定時に空白や0を返すことがある）
  blankValues: ['(not set)', '0'],
  rowLimit: 10000,      // 1リクエストの行数
  maxRows: 10000        // 1サイトあたりの最大行数（JSON肥大化の防止）
};

/** 確認用：各プロパティのカスタムディメンション（apiName／表示名／説明）をログに出す */
function debug_listCustomDims(){
  for (var site in CONFIG.ga4Properties){
    try {
      var md = AnalyticsData.Properties.getMetadata('properties/'+CONFIG.ga4Properties[site]+'/metadata');
      var custom = (md.dimensions||[]).filter(function(d){ return d.customDefinition; });
      if (!custom.length) Logger.log('%s\t（カスタムディメンションなし）', site);
      custom.forEach(function(d){ Logger.log('%s\t%s\t%s\t%s', site, d.apiName, d.uiName, d.description||''); });
      var dims = resolveDocodocoDims_(CONFIG.ga4Properties[site]);
      Logger.log('%s\t→ 使用するディメンション: %s', site, JSON.stringify(dims));
    } catch(err){ Logger.log('metadata失敗 %s: %s', site, err); }
  }
}

/** プロパティごとに使用するディメンション名を決める（未登録の項目は空文字） */
function resolveDocodocoDims_(propertyId){
  var md = AnalyticsData.Properties.getMetadata('properties/'+propertyId+'/metadata');
  var custom = (md.dimensions||[]).filter(function(d){ return d.customDefinition; });
  var exists = {}; custom.forEach(function(d){ exists[d.apiName] = true; });
  var dims = {};
  for (var k in DOCODOCO_CONFIG.dims){
    var spec = DOCODOCO_CONFIG.dims[k], found = '';
    (spec.api||[]).some(function(a){ if (exists[a]) { found = a; return true; } return false; });
    if (!found){
      var cand = custom.filter(function(d){ return spec.src.test(d.description||''); });
      (spec.ui||[]).some(function(re){
        var hit = cand.filter(function(d){ return re.test(d.uiName||''); })[0];
        if (hit) { found = hit.apiName; return true; }
        return false;
      });
    }
    dims[k] = found;
  }
  return dims;
}

/** 値が空・(not set) でない条件 */
function notBlank_(dim){
  return { notExpression:{ orGroup:{ expressions:[
    { filter:{ fieldName:dim, stringFilter:{ matchType:'EXACT', value:'(not set)' } } },
    { filter:{ fieldName:dim, emptyFilter:{} } }
  ] } } };
}

/** メイン：企業×閲覧日×閲覧ページ ＋ 業種×年月 */
function fetchDocodoco_(propertyId, end){
  if (!DOCODOCO_CONFIG.enabled) return null;
  var prop = 'properties/'+propertyId;
  var dims = resolveDocodocoDims_(propertyId);
  if (!dims.org && !dims.usCompany){
    return { error:'組織名（どこどこJP）・社名（uSonar）のカスタムディメンションが見つかりません。debug_listCustomDims で確認してください', dims:dims };
  }
  var start = isoDaysAgo(DOCODOCO_CONFIG.lookbackDays);

  // どこどこJPの組織名 または uSonarの社名 のどちらかが取れている行だけを対象にする
  var conds = [];
  if (dims.org) conds.push(notBlank_(dims.org));
  if (dims.usCompany) conds.push(notBlank_(dims.usCompany));
  var companyFilter = conds.length === 1 ? conds[0] : { orGroup:{ expressions: conds } };

  // (1) 企業×閲覧日×閲覧ページ（GA4 APIのディメンション上限9個以内：最大8個）
  var keys = ['org','industry','sales','usCompany','houjin','capital'];
  var dimList = [{name:'date'}], idx = {};
  keys.forEach(function(k){ if (dims[k]) { idx[k] = dimList.length; dimList.push({name:dims[k]}); } });
  var pathIdx = dimList.length; dimList.push({name:'unifiedPagePathScreen'});

  var visits = [], offset = 0, total = 0;
  do {
    var r = AnalyticsData.Properties.runReport({
      dateRanges:[{startDate:start, endDate:end}],
      dimensions: dimList,
      metrics:[{name:'screenPageViews'}, {name:'sessions'}],
      dimensionFilter: companyFilter,
      orderBys:[{dimension:{dimensionName:'date'}, desc:true}],
      limit: DOCODOCO_CONFIG.rowLimit,
      offset: offset
    }, prop);
    total = r.rowCount || 0;
    (r.rows||[]).forEach(function(row){
      var dv = row.dimensionValues, d = dv[0].value;
      var v = {
        date: d.slice(0,4)+'-'+d.slice(4,6)+'-'+d.slice(6,8),
        org:'', industry:'', sales:'', usCompany:'', houjin:'', capital:'',
        path: dv[pathIdx].value,
        views: +row.metricValues[0].value,
        sessions: +row.metricValues[1].value
      };
      keys.forEach(function(k){
        if (idx[k] !== undefined){
          var val = String(dv[idx[k]].value || '').trim();
          v[k] = (DOCODOCO_CONFIG.blankValues.indexOf(val) >= 0) ? '' : val;
        }
      });
      if (DOCODOCO_CONFIG.excludeOrgRegex && v.org && DOCODOCO_CONFIG.excludeOrgRegex.test(v.org) && !v.usCompany) return;
      if (v.houjin && DOCODOCO_CONFIG.excludeHoujin.indexOf(v.houjin) >= 0) return;
      if (!v.org && !v.usCompany && !v.houjin) return;   // 空白だけの行は除外
      visits.push(v);
    });
    offset += DOCODOCO_CONFIG.rowLimit;
  } while (offset < total && offset < DOCODOCO_CONFIG.maxRows);

  // (2) どこどこJPの業種大分類×年月（期間はGA4本体と同じ CONFIG.lookbackDays）
  var industryMonthly = [];
  if (dims.industry){
    try {
      var rI = AnalyticsData.Properties.runReport({
        dateRanges:[{startDate:isoDaysAgo(CONFIG.lookbackDays), endDate:end}],
        dimensions:[{name:'yearMonth'}, {name:dims.industry}],
        metrics:[{name:'sessions'}],
        dimensionFilter: notBlank_(dims.industry),
        limit: 5000
      }, prop);
      industryMonthly = (rI.rows||[]).map(function(row){
        var ym = row.dimensionValues[0].value;
        return { ym: ym.slice(0,4)+'/'+ym.slice(4,6), industry: row.dimensionValues[1].value,
                 sessions: +row.metricValues[0].value };
      });
    } catch(err){ Logger.log('docodoco industryMonthly失敗 %s: %s', propertyId, err); }
  }

  return { dims:dims, period:start+'〜'+end, rowCount:total, truncated: total > offset,
           visits:visits, industryMonthly:industryMonthly };
}
// ▲▲▲ 訪問企業データの追加 ここまで ▲▲▲

// ===================================================================
//  ▼▼▼ エンゲージメントの追加（2026-11 追加） ▼▼▼
//  目的：DSP広告の判定項目「LPの直帰率の低下」「LPの滞在時間の延長」に使うデータを、GA4から取る。
//  ※「エンゲージメント」「直帰率」は GA4 の指標で、Search Console（GSC）にはありません。
//    GA4の直帰率＝エンゲージしなかったセッションの割合＝1−エンゲージメント率。
//    「エンゲージしたセッション」＝10秒以上の滞在、またはコンバージョン、または2ページ以上の閲覧
//
//  出力：payload.ga4[サイト].engagement = {
//    period:         取得期間
//    monthly:        [{ym:'2026/10', sessions, engagedSessions, engagementRate, bounceRate, engagementSec, avgEngagementSec, viewsPerSession}]   … サイト全体の月次
//    channelMonthly: [{ym, ch:'Display', ...同じ指標}]                                                                                        … チャネル別の月次
//    dspMonthly:     [{ym, sourceMedium:'admatrix / display', ...}]                                                                           … DSP流入の月次（DSP対象サイトのみ）
//    dspLanding:     [{ym, page:'/', sourceMedium, ...}]                                                                                      … DSP流入のLP別の月次（DSP対象サイトのみ）
//    deviceMonthly:  [{ym, device:'desktop'|'mobile'|'tablet', ...同じ指標}]                                                                  … デバイス別（PC／SP）の月次（サイト全体）
//    dspDevice:      [{ym, device, sourceMedium, ...}]                                                                                        … DSP流入のデバイス別（PC／SP）の月次（DSP対象サイトのみ）
//  }
//  指標：sessions＝セッション数／engagedSessions＝エンゲージしたセッション数／engagementRate・bounceRate＝0〜1の割合
//        engagementSec＝エンゲージメント時間の合計（秒）／avgEngagementSec＝1セッションあたり（秒）／viewsPerSession＝1セッションあたりの閲覧ページ数
//  メモ：「次ページ遷移率」（月次レポートの約15%）は、GA4のAPIに直接の指標がありません。当面は viewsPerSession（回遊）と
//        直帰率で代替します（セッション単位で「2ページ以上見た割合」が必要なら、GA4の探索レポートか BigQuery エクスポートが必要）。
//  確認用：debug_engagement（ログに出す）／run_buildEngagementJson（単独の確認用ファイル dsp_dashboard_engagement.json を出す）
// ===================================================================
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
function engMetrics_(mv){
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
function engYm_(v){ return v.slice(0, 4) + '/' + v.slice(4, 6); }

/** runReport をページ送りで全件取得する */
function engRun_(prop, req){
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
function engIsDspProperty_(propertyId){
  return ENGAGEMENT_CONFIG.dspSites.some(function(site){
    return String(CONFIG.ga4Properties[site]) === String(propertyId);
  });
}

/** メイン：fetchGa4_ から呼ぶ（start/end は fetchGa4_ と同じ期間） */
function fetchEngagement_(propertyId, start, end){
  if (!ENGAGEMENT_CONFIG.enabled) return null;
  var prop = 'properties/' + propertyId;
  var range = [{ startDate: start, endDate: end }];
  var out = { period: start + '〜' + end, monthly: [], channelMonthly: [], dspMonthly: [], dspLanding: [], deviceMonthly: [], dspDevice: [] };

  // (1) サイト全体の月次
  out.monthly = engRun_(prop, {
    dateRanges: range, dimensions: [{ name: 'yearMonth' }], metrics: ENG_METRICS,
    orderBys: [{ dimension: { dimensionName: 'yearMonth' } }]
  }).map(function(row){
    var o = engMetrics_(row.metricValues); o.ym = engYm_(row.dimensionValues[0].value); return o;
  });

  // (2) チャネル別の月次
  out.channelMonthly = engRun_(prop, {
    dateRanges: range, dimensions: [{ name: 'yearMonth' }, { name: 'sessionDefaultChannelGroup' }], metrics: ENG_METRICS
  }).map(function(row){
    var o = engMetrics_(row.metricValues);
    o.ym = engYm_(row.dimensionValues[0].value); o.ch = row.dimensionValues[1].value; return o;
  });

  // (2b) デバイス別（PC＝desktop／SP＝mobile／tablet）の月次（サイト全体）
  try {
    out.deviceMonthly = engRun_(prop, {
      dateRanges: range, dimensions: [{ name: 'yearMonth' }, { name: 'deviceCategory' }], metrics: ENG_METRICS
    }).map(function(row){
      var o = engMetrics_(row.metricValues);
      o.ym = engYm_(row.dimensionValues[0].value); o.device = row.dimensionValues[1].value; return o;
    });
  } catch(err){ Logger.log('GA4 deviceMonthly失敗 %s: %s', propertyId, err); }

  // (3)(4)(5) DSP流入（DSP対象サイトのみ）
  if (engIsDspProperty_(propertyId)){
    var dspFilter = { filter: { fieldName: 'sessionSourceMedium',
      stringFilter: { matchType: 'PARTIAL_REGEXP', value: ENGAGEMENT_CONFIG.dspSourceMediumRegex, caseSensitive: false } } };
    try {
      out.dspMonthly = engRun_(prop, {
        dateRanges: range, dimensions: [{ name: 'yearMonth' }, { name: 'sessionSourceMedium' }], metrics: ENG_METRICS,
        dimensionFilter: dspFilter
      }).map(function(row){
        var o = engMetrics_(row.metricValues);
        o.ym = engYm_(row.dimensionValues[0].value); o.sourceMedium = row.dimensionValues[1].value; return o;
      });
    } catch(err){ Logger.log('GA4 dspMonthly失敗 %s: %s', propertyId, err); }
    try {
      out.dspLanding = engRun_(prop, {
        dateRanges: range,
        dimensions: [{ name: 'yearMonth' }, { name: 'landingPage' }, { name: 'sessionSourceMedium' }],
        metrics: ENG_METRICS, dimensionFilter: dspFilter
      }).map(function(row){
        var o = engMetrics_(row.metricValues);
        o.ym = engYm_(row.dimensionValues[0].value); o.page = row.dimensionValues[1].value;
        o.sourceMedium = row.dimensionValues[2].value; return o;
      });
    } catch(err){ Logger.log('GA4 dspLanding失敗 %s: %s', propertyId, err); }
    try {
      out.dspDevice = engRun_(prop, {
        dateRanges: range,
        dimensions: [{ name: 'yearMonth' }, { name: 'deviceCategory' }, { name: 'sessionSourceMedium' }],
        metrics: ENG_METRICS, dimensionFilter: dspFilter
      }).map(function(row){
        var o = engMetrics_(row.metricValues);
        o.ym = engYm_(row.dimensionValues[0].value); o.device = row.dimensionValues[1].value;
        o.sourceMedium = row.dimensionValues[2].value; return o;
      });
    } catch(err){ Logger.log('GA4 dspDevice失敗 %s: %s', propertyId, err); }
  }
  return out;
}

/** 確認用：DSP対象サイトの直近月の値をログに出す（先にこれを実行して認証・値を確認する） */
function debug_engagement(){
  var start = isoDaysAgo(120), end = isoToday();
  ENGAGEMENT_CONFIG.dspSites.forEach(function(site){
    var pid = CONFIG.ga4Properties[site];
    var e = fetchEngagement_(pid, start, end);
    Logger.log('== %s（property %s）期間 %s', site, pid, e.period);
    e.monthly.forEach(function(m){
      Logger.log('  サイト全体 %s  セッション%s  エンゲージ率%s%%  直帰率%s%%  平均エンゲージ%s秒  閲覧/セッション%s',
        m.ym, m.sessions, (m.engagementRate * 100).toFixed(1), (m.bounceRate * 100).toFixed(1), m.avgEngagementSec.toFixed(1), m.viewsPerSession.toFixed(2));
    });
    e.dspMonthly.forEach(function(m){
      Logger.log('  DSP %s  %s  セッション%s  エンゲージ率%s%%  直帰率%s%%  平均エンゲージ%s秒  閲覧/セッション%s',
        m.ym, m.sourceMedium, m.sessions, (m.engagementRate * 100).toFixed(1), (m.bounceRate * 100).toFixed(1), m.avgEngagementSec.toFixed(1), m.viewsPerSession.toFixed(2));
    });
    e.dspDevice.filter(function(m){ return m.sourceMedium === 'admatrix / display' && m.ym >= '2026/09'; })
      .sort(function(a, b){ return a.ym < b.ym ? -1 : a.ym > b.ym ? 1 : (a.device < b.device ? -1 : 1); }).forEach(function(m){
      Logger.log('  DSP端末別 %s  %s  セッション%s  直帰率%s%%  平均エンゲージ%s秒  閲覧/セッション%s',
        m.ym, m.device, m.sessions, (m.bounceRate * 100).toFixed(1), m.avgEngagementSec.toFixed(1), m.viewsPerSession.toFixed(2));
    });
    var top = e.dspLanding.slice().sort(function(a, b){ return b.sessions - a.sessions; }).slice(0, 5);
    top.forEach(function(m){
      Logger.log('  DSPのLP上位 %s  %s  セッション%s  直帰率%s%%', m.ym, m.page, m.sessions, (m.bounceRate * 100).toFixed(1));
    });
  });
}

/** 単独で確認用ファイル dsp_dashboard_engagement.json をドライブに出力する（既存の run_buildJson とは独立） */
function run_buildEngagementJson(){
  var start = isoDaysAgo(CONFIG.lookbackDays), end = isoToday();
  var out = { type: 'engagement', generated: new Date().toISOString(), ga4: {} };
  for (var site in CONFIG.ga4Properties){
    try { out.ga4[site] = fetchEngagement_(CONFIG.ga4Properties[site], start, end); }
    catch(err){ Logger.log('GA4 engagement失敗 %s: %s', site, err); }
  }
  var name = 'dsp_dashboard_engagement.json', text = JSON.stringify(out);
  var folder = CONFIG.driveFolderId ? DriveApp.getFolderById(CONFIG.driveFolderId) : DriveApp.getRootFolder();
  var it = folder.getFilesByName(name);
  if (it.hasNext()) { it.next().setContent(text); } else { folder.createFile(name, text, MimeType.PLAIN_TEXT); }
  Logger.log('生成完了: %s サイト → %s', Object.keys(out.ga4).length, name);
}
// ▲▲▲ エンゲージメントの追加 ここまで ▲▲▲
