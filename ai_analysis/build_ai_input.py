# -*- coding: utf-8 -*-
"""AI分析に渡す集計表（個人情報なし）を作る。

使い方（Python は Anaconda を直接指定。読み込むファイルがあるため -I を付ける）:
  C:\\work\\anaconda_install\\python.exe -I ai_analysis\\build_ai_input.py IS

- 入力ファイルは「コピーしてから読む」（開いたまま・書き込みはしない）。コピー先は一時フォルダ。
- 出力：ai_analysis/input/<事業>_AI分析用集計表_<日付>.xlsx
- 個人情報（氏名・メール・会社名・電話・受注企業名）は出力しない。実績.xlsm は件数だけ集計する。
- 2026-11-25 に IS 版を作成。他事業（SCH など）は BIZ_CONF と「月次レポート所見」を足せば同じ形で作れる。
"""
import sys, os, re, csv, json, glob, shutil, tempfile, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BIZ = sys.argv[1] if len(sys.argv) > 1 else 'IS'
TODAY = datetime.date.today().isoformat()

DOC = r'C:\Users\612316\Box\マーケティンググループ\デジタルマーケティングDiv\マーケティング本部\社内ドキュメント'
GEO = r'C:\Users\612316\Documents\GitHub\GEO\monitoring\data'
SHISAKU = r'C:\Users\612316\Box\【外部】マーケティング本部\施策管理表.xlsx'
HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(HERE, 'input')

BIZ_CONF = {
    'IS': dict(
        jisseki=os.path.join(DOC, 'マーケティング本部｜IS｜実績.xlsm'),
        geo_domain='インサイトセールス', site='solution.onward-cd.co.jp',
        dsp_group='IS', ga_site='solution.onward-cd.co.jp', sheet='IS'),
}
C = BIZ_CONF[BIZ]

tmp = tempfile.mkdtemp(prefix='ai_input_')


def cp(src, name):
    dst = os.path.join(tmp, name)
    shutil.copy2(src, dst)
    return dst


# ---------------------------------------------------------------- 出力の道具
wb = openpyxl.Workbook()
wb.remove(wb.active)
HDR = PatternFill('solid', fgColor='1F3A5F')
SUB = PatternFill('solid', fgColor='E8EEF7')


def sheet(title, header, rows, widths=None, note=None, wrap_cols=()):
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(1, 1, note).font = Font(italic=True, color='555555')
        ws.cell(1, 1).alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(header)))
        ws.row_dimensions[1].height = 48
        r0 = 2
    for j, h in enumerate(header, 1):
        c = ws.cell(r0, j, h)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for i, row in enumerate(rows, r0 + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.alignment = Alignment(wrap_text=(j - 1) in wrap_cols or isinstance(v, str) and len(v) > 40, vertical='top')
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths[j - 1] if widths and j - 1 < len(widths) else 14)
    ws.freeze_panes = ws.cell(r0 + 1, 1)
    return ws


def pct(a, b):
    return round(a / b, 4) if b else None


# ---------------------------------------------------------------- ROI
roi = json.load(open(os.path.join(DOC, 'marketing_roi_2026_2026-11-09.json'), encoding='utf-8'))
act = roi['actuals'][BIZ]
tg = roi['targets'][BIZ]['2026']
roi_rows = []
for ym in sorted(act):
    v = act[ym]
    y, m = int(ym[:4]), int(ym[5:])
    if (y, m) > (2026, 10):
        continue  # 11月以降は未入力（0ではない）
    t = roi['targets'][BIZ].get(str(y), {})

    def tv(k):
        x = t.get(k)
        return (x[m - 1] if isinstance(x, list) else x) if x is not None else None
    roi_rows.append([ym, v.get('sessions'), tv('monthly_session'), v.get('cv'), tv('monthly_mql'), v.get('sal'),
                     v.get('sql'), tv('monthly_sql'), v.get('jucyu'), v.get('content_count'), v.get('cost_content'),
                     pct(v.get('cv') or 0, v.get('sessions') or 0)])
months = [f'2026-{m:02d}' for m in range(3, 11)]
S = sum(act[m]['sessions'] for m in months); Cv = sum(act[m]['cv'] for m in months)
Sq = sum(act[m]['sql'] for m in months); Sa = sum(act[m]['sal'] for m in months)
tS = sum(tg['monthly_session'][int(m[-2:]) - 1] for m in months)
tM = sum(tg['monthly_mql'][int(m[-2:]) - 1] for m in months)
tQ = sum(tg['monthly_sql'][int(m[-2:]) - 1] for m in months)
roi_note = ('出どころ：社内ドキュメント marketing_roi_2026_2026-11-09.json。MQL＝「cv」の項目。'
            '2026-11以降は未入力（実績0ではない）ため除外。月次レポートの数字はこのROIダッシュボードの機能で作成しているため、'
            '両者が一致するのは当然で、独立した裏付けにはならない（根拠は1つとして数える）。')
sheet('01_ROI_月次', ['年月', 'セッション', 'セッション目標', 'MQL(cv)', 'MQL目標', 'SAL', 'SQL', 'SQL目標', '受注',
                      'コンテンツ本数', 'コンテンツ費用', 'MQL÷セッション'], roi_rows,
      [10, 11, 13, 9, 9, 7, 7, 9, 7, 12, 14, 14], roi_note)
ws = wb['01_ROI_月次']
r = ws.max_row + 2
ws.cell(r, 1, '2026年度 3〜10月累計').font = Font(bold=True)
for i, (k, v) in enumerate([
        ('セッション', f'{S:,}／目標{tS:,.0f}（{S / tS * 100:.0f}%）'),
        ('MQL', f'{Cv}／目標{tM:.0f}（{Cv / tM * 100:.0f}%）'),
        ('SAL', f'{Sa}'), ('SQL', f'{Sq}／目標{tQ:.0f}（{Sq / tQ * 100:.0f}%）'),
        ('SQL÷SAL', f'{Sq / max(Sa, 1) * 100:.0f}%'), ('SQL÷MQL', f'{Sq / max(Cv, 1) * 100:.0f}%')], 1):
    ws.cell(r + i, 1, k); ws.cell(r + i, 2, v)

# ---------------------------------------------------------------- 月次レポート所見（人が書いた所見。社名・個人名は除く）
rep_rows = [
    ['サマリー', 'MAL・MQLは未達、SQLは3件で達成。', 'MQL年累計達成率62%／SQL 112.5%', 'p.39'],
    ['アクセス', 'DSP広告（Display）の予算は通常に戻したが、流入数自体は好調。（配信会社の管理画面が閲覧できないため詳細は振り返りMTGの資料で確認）', 'セッション3,381・MQL5', 'p.40'],
    ['CV', 'SQL3件のうち2件は外部団体の研究会関連のSALからの引き上がり。9月にMQLを多数獲得できたWP7は「担当者のペインポイントを突いたテーマ」。同種のコンテンツ拡充で接点を創出できる可能性。', 'SAL0・SQL3・受注0／資料DL5・問合せ0・WP0', 'p.41'],
    ['コラム', '10月公開2本（10/7・10/16）。表示回数・セッションは前月比で上昇し、新作2本とメルマガ誘導で表示回数は約2倍。1位のコラムは閲覧数は多いが読了率が著しく低い。上位1・2位はメルマガ流入。Organic流入は福袋コラムが依然トップ。', '公開本数：6月以外は月1本ペース、3・5・10月は2本', 'p.42'],
    ['事例', '10月公開1本（10/28）。今期の公開は3月2本・10月1本。表示回数は前月比で約200増。事例は本文が短く読了率（90%）が低くなりがちなので、平均エンゲージメント時間で評価するのが適切。新規公開の事例は閲覧数は少ないがエンゲージメントが高い傾向。「アップサイクル」「理念の体現」の文脈が刺さっている可能性。', '', 'p.43'],
    ['メルマガNo.53', '開封率は高いがクリック率が低く、評価△。読者の約75%が「ちら読み」。本文にデータやグラフを載せすぎ、本文だけで満足されて詳細ページへ進む動機が薄れた、という仮説。改善案：データを1つに絞り、「続きはコラムで」へ誘導／コラムで得られるメリットを強調／ファーストビュー付近にCTAを置く。', '送信5,184・開封率17.73%・クリック率0.31%（15件）・CV0・ちら読み75.5%（読了14.8%）', 'p.44'],
    ['メルマガNo.52', '開封率は約16%だがクリック率が低く、評価△。冒頭がテキスト主体の課題提起で、メリットが伝わる前に離脱されているという仮説。改善案：冒頭のお悩みを視覚的なチェックリストに／CTA文言を、クリック後のベネフィットが分かる表現に。', '送信5,159・開封率15.91%・クリック率0.45%（19件）・CV0・ちら読み75.5%（読了16.5%）', 'p.45'],
    ['DSP広告（9月実績）', '8月の配信停止による残予算を9月に追加投入。7月と比べ imp 165%・クリック約290%・CPC 96円→63円。SPはCPM（表示単価）が高騰したがクリック率が大幅に上昇。PCは全素材でCTRが低下し始めており、クリエイティブの差し替えを検討。', '総imp 3,415,315・クリック6,001', 'p.46'],
    ['課題と施策（既存）', '施策No.178：業界業種のグルーピング＋業種別メルマガ（グルーピング完了、業種別シナリオメールは計画中）。No.179：他事業のハウスリストへの配信は検討優先順位が低い。No.180：WP起点メルマガ（WPの改修に着手。WPの本数が増えたためWP起点の内製メルマガを検討）。メルマガ経由の資料導線に限定しすぎると資料閲覧で行動が終わりがちなので、コラム・事例など回遊するページへの流入も設計する。', '', 'p.47'],
    ['示唆①（人が記載）', '担当者のペインポイントを突いたコンテンツの拡充による接点創出。例：環境配慮型グッズ製作で失敗しないためのグリーンウォッシュ回避／ブランド価値を損なわない「サステナブル販促」／販促担当者が知っておくべき景品表示法・著作権・商標権の基礎知識／急なイベントを成功に導く逆算型スケジュール管理／周年記念品・社内グッズの設計図。', '※資料に「検討の上、施策管理表への起票・実施」と記載', 'p.48'],
    ['示唆②（人が記載）', 'メルマガの訴求とコラム内容のマッチ度、見出しの作り方、全体のボリュームの見直し。メルマガサムネイルとコラム画像を連動／「続きを読みたい」と思わせる構成（メルマガクリック率の改善）／ロング版・ショート版のABテスト。', '', 'p.48'],
    ['示唆③（人が記載）', 'コラム読了率の向上。記事冒頭に「この記事で得られる3つのメリット」を箇条書きで明示／見出し（目次）の短文化。エンタープライズの多忙なマーケター向けに「拾い読み」ができる構成にする。', '', 'p.48'],
    ['年間戦略と課題', '年間戦略：SQL②の件数増加・良質化（認知拡大・MAL増加）。MQL以前の課題：①事業認知の不足（Organic流入が目標に対し最も進捗率が低く、指名検索のImpも少なくクリック0＝9月）②MALの不足（9月はDisplay増額とEmailで目標達成、Organicはほぼ変わらず）。検討したいこと：非指名検索KWによる初回流入の創出（短期＝リスティングの再導入、長期＝非指名KWのコンテンツマーケ）／DSP広告の改善（サイトデザインに近いクリエイティブ、非指名型の課題訴求型コピーのテスト、流入後の回遊改善）／強いヘッドラインを持つLPが必要（セールスプロモーションのLPを検討）／まずDSP広告の改善から実施（業務委託デザイナーの調整後）。', 'DSP広告からTOPページ流入後の次ページ遷移率：約15%', 'p.49'],
]
sheet('02_月次レポート所見', ['区分', '内容（社名・個人名を除いて要約）', '数字', '出典（FLUED｜2026｜月次定例レポート 10月号のスライド）'], rep_rows,
      [18, 110, 42, 18],
      '出どころ：FLUED様の月次定例レポート（2026年10月号・IS）。人が書いた所見・示唆を含む。SAL/SQL企業名・事例企業名は除いて要約した。数字（セッション・MQL等）はROIダッシュボードの出力と同じもの。',
      wrap_cols=(1, 2))

# ---------------------------------------------------------------- GEO（設問別）
ins = json.load(open(os.path.join(GEO, 'reports', 'insights_20261104_133016.json'), encoding='utf-8'))
res = list(csv.DictReader(open(os.path.join(GEO, 'results', 'results_20261104_133016_r1.csv'), encoding='utf-8-sig')))
geo_rows = []
for r in res:
    if r['domain_label'] != C['geo_domain']:
        continue
    geo_rows.append([r['question_id'], r['type_label'], r['stakeholder_label'], r['question'],
                     '出た' if str(r['mention_detected']).strip().lower() in ('true', '1', 'yes') else '出なかった',
                     r['mention_position'], r['entities_found'], r['urls_found'], r['specificity_tier'], r['model_id'], r['answer']])
sheet('03_GEO_設問別', ['設問ID', '質問タイプ', '想定読者', '質問', '自社の出現', '出現位置', '回答に出た企業名', '回答に出たURL', '具体度', 'モデル', '回答（全文）'],
      geo_rows, [10, 14, 12, 50, 10, 8, 40, 30, 8, 14, 120],
      f'出どころ：GEO insights/results {ins["generated_at"]}（2026-11-04実行）。{C["geo_domain"]}の全{len(geo_rows)}問。全問が1モデル（claude-sonnet-5）の結果。'
      '競合の社名はAIに渡してよい（中島さん確認済み）。質問の件数が少なく1モデルのため、GEO由来の候補は根拠を「弱い」とする。', wrap_cols=(3, 6, 10))

# 実行履歴（領域×タイプ別の出現）
hist = []
for f in sorted(glob.glob(os.path.join(GEO, 'results', 'results_*_r1.csv'))):
    rows = list(csv.DictReader(open(f, encoding='utf-8-sig')))
    tot = collections.Counter(); det = collections.Counter()
    for r in rows:
        if r['domain_label'] != C['geo_domain']:
            continue
        tot[r['type_label']] += 1
        if str(r['mention_detected']).strip().lower() in ('true', '1', 'yes'):
            det[r['type_label']] += 1
    _p = os.path.basename(f).split('_')
    run = _p[1] + '_' + _p[2]
    hist.append([run, sum(det.values()), sum(tot.values())] + [f'{det[t]}/{tot[t]}' for t in sorted({r['type_label'] for r in res if r['domain_label'] == C['geo_domain']})])
types = sorted({r['type_label'] for r in res if r['domain_label'] == C['geo_domain']})
sheet('03b_GEO_推移', ['実行日', '出現', '質問数'] + types, hist, [12, 7, 8] + [14] * len(types),
      f'{C["geo_domain"]}の出現の推移（質問タイプ別の「出現／質問数」）。')

# 定型文（GEOの改善案）
pb = [[a['type'], f'{a["miss"]}/{a["total"]}', a['priority'], a['playbook'], '、'.join(a.get('cover_attributes', [])), '、'.join(a.get('beating_competitors', []))]
      for a in ins['actions'] if a['domain'] == C['geo_domain']]
sheet('04_GEO_定型文', ['質問タイプ', '出なかった／質問数', '優先度', '定型文（現状のまま）', '回答で重視された観点', '想起された競合'], pb,
      [14, 14, 8, 80, 50, 40],
      '定型文はいったん現状の表現のまま（中島さん確認済み）。施策候補の内容は、この定型文ではなく、上の「設問別」の結果からAIが分析した示唆にする。',
      wrap_cols=(3, 4, 5))

# ---------------------------------------------------------------- DSP広告・サイト（最新ダッシュボード 2026-11-13）
html = open(cp(os.path.join(DOC, 'DSP広告分析ダッシュボード.html'), 'dsp.html'), encoding='utf-8').read()
i = html.index('const SEED={"version"') + len('const SEED=')
seed, _ = json.JSONDecoder().raw_decode(html, i)
dsp_rows = []
bud = seed.get('budget', {})
for ym in sorted(seed['dsp']['trend']):
    v = seed['dsp']['trend'][ym].get(C['dsp_group'])
    if not v:
        continue
    dsp_rows.append([ym, v['imp'], v['clk'], pct(v['clk'], v['imp']), round(v['cpc'], 1) if v.get('cpc') else None, v['cost'],
                     bud.get(ym, {}).get(C['dsp_group']), v.get('cv'), v.get('vtcv')])
sheet('05_DSP広告_月次', ['年月', 'imp', 'クリック', 'CTR', 'CPC(円)', '配信費用(円)', '月予算(円)', 'CV', 'ビュースルーCV'], dsp_rows,
      [10, 12, 10, 9, 9, 13, 12, 6, 12],
      f'出どころ：DSP広告分析ダッシュボード（{seed["updated"][:10]}更新・2026年10月分まで取込済み）。DSP広告はSCHとISの2事業で実施。目的は認知アップ・エンゲージメントアップ。'
      '月次レポート9月号の配信実績（imp 3,415,315・クリック6,001）と一致。')
bm = seed['dsp']['byMonth']
g_rows = []
for ym in sorted(bm):
    for g in bm[ym].get('groups', []):
        if f'_{C["dsp_group"]}_' in g['name']:
            g_rows.append([ym, g['name'], g['imp'], g['clk'], pct(g['clk'], g['imp']), g['cost']])
sheet('05b_DSP_PC・SP別', ['年月', '広告グループ', 'imp', 'クリック', 'CTR', '費用(円)'], g_rows, [10, 20, 12, 10, 9, 12], 'PC／SP別の月次。')
c_rows = []
for ym in sorted(bm):
    for c in bm[ym].get('creatives', []):
        if f'_{C["dsp_group"]}_' in c['name']:
            c_rows.append([ym, c['name'], c['size'], c['device'], c['imp'], c['clk'], pct(c['clk'], c['imp']), c.get('cost')])
sheet('05c_DSP_クリエイティブ別', ['年月', 'クリエイティブ名', 'サイズ', '端末', 'imp', 'クリック', 'CTR', '費用(円)'], c_rows,
      [10, 36, 10, 6, 12, 10, 9, 10], '素材別の月次（IS分）。配信先企業の名称は含めない。')

ga = seed['ga4'][C['ga_site']]
ch = collections.defaultdict(dict)
for x in ga['report']['channelMonthly']:
    ch[x['ym']][x['ch']] = x['sessions']
chs = sorted({k for v in ch.values() for k in v})
ch_rows = [[ym] + [ch[ym].get(k, 0) for k in chs] for ym in sorted(ch)]
sheet('06_サイトGA4_チャネル月次', ['年月'] + chs, ch_rows, [10] + [12] * len(chs),
      f'出どころ：同ダッシュボード内のGA4（{C["ga_site"]}・取得期間{ga["period"]}）。チャネル別セッション数の月次。')
lp_rows = []
for m in ga['report']['months']:
    if not m.get('complete'):
        continue
    for l in m.get('landing', []):
        if 'admatrix' in l['sourceMedium']:
            lp_rows.append([m['ym'], l['page'], l['sourceMedium'], l['sessions'], l['activeUsers'], l['newUsers'], l['engagementSec'],
                            round(l['engagementSec'] / l['sessions'], 1) if l['sessions'] else None])
sheet('06b_DSP流入のLP別', ['年月', 'ランディングページ', '流入元/メディア', 'セッション', 'アクティブUU', '新規UU', 'エンゲージメント秒（合計）', '1セッションあたり(秒)'], lp_rows,
      [10, 40, 20, 11, 12, 10, 18, 16],
      'DSP広告（admatrix／display）経由のランディングページ別の指標。月次レポートの「次ページ遷移率約15%」を補う。'
      '※LPの直帰率はこのダッシュボードのデータにない（GA4の直帰率が未取得）。滞在時間の代わりにエンゲージメント時間を使う。')
pg_rows = []
for m in ga['report']['months']:
    if not m.get('complete'):
        continue
    for p in sorted(m['pages'], key=lambda x: -x['views'])[:15]:
        pg_rows.append([m['ym'], p['path'], p['views'], p['sessions'], round(p['viewsPerSession'], 2),
                        round(p['engagementSec'] / p['sessions'], 1) if p['sessions'] else None, p['scrollUsers'], p['users']])
sheet('06c_ページ別上位', ['年月', 'ページ', '表示回数', 'セッション', '回遊(ページ/セッション)', '1セッションあたり滞在(秒)', 'スクロールUU', 'UU'], pg_rows,
      [10, 56, 10, 11, 14, 16, 12, 8], '表示回数上位15ページ（月ごと）。')
gsc = seed['gscsearch'].get(C['dsp_group'], {})
gm = [[x['ym'], x['imp'], x['clk'], round(x['ctr'], 4), round(x['pos'], 1)] for x in gsc.get('monthly', [])]
sheet('06d_GSC_月次', ['年月', '表示回数', 'クリック', 'CTR', '平均掲載順位'], gm, [10, 12, 10, 9, 12],
      f'出どころ：同ダッシュボード内のGSC（{gsc.get("period")}）。指名検索の比率：{gsc.get("brandShare", 0) * 100:.1f}%。')
q_rows = [[x['q'], x['clk'], x['imp'], round(x['ctr'], 4), round(x['pos'], 1), x.get('cls'), '指名' if x.get('brand') else '非指名'] for x in gsc.get('topQueries', [])]
sheet('06e_GSC_検索クエリ上位', ['クエリ', 'クリック', '表示回数', 'CTR', '平均掲載順位', '分類', '指名/非指名'], q_rows, [40, 9, 10, 9, 12, 10, 10], '検索クエリ上位50。')

# ---------------------------------------------------------------- 実績ファイル（CVリスト）：件数だけ集計
jf = cp(C['jisseki'], 'jisseki.xlsm')
wbj = openpyxl.load_workbook(jf, read_only=True, data_only=True)
wsj = wbj['流入チャネル別顧客一覧']
rows = list(wsj.iter_rows(values_only=True))
head = [str(h or '').strip() for h in rows[0]]
ix = {h: i for i, h in enumerate(head)}


def col(name):
    for h, i in ix.items():
        if h.startswith(name):
            return i
    raise KeyError(name)


ct, cc, cm = col('CVフラグ'), col('流入チャネル'), col('CV年月')
flags = {'SAL': col('SAL'), 'SQL1': col('SQL1'), 'SQL2': col('SQL2'), '提案': col('提案した'), '受注': col('受注'), '辞退': col('辞退'), '失注': col('失注')}
by_m_t = collections.Counter(); by_m_c = collections.Counter(); fun_t = collections.defaultdict(collections.Counter); fun_c = collections.defaultdict(collections.Counter)
n = 0
fun_l = collections.defaultdict(collections.Counter)


def grp(label):
    """CVの元（ラベル）を大きく分ける。メルマガのCTA名は細かいので、別シートに一覧で残す。"""
    if label.startswith(BIZ + '｜') and 'メルマガ' in label:
        return 'メルマガ経由（CTA）'
    if label.startswith('入力画面') or label.lower().startswith('input screen'):
        return 'フォーム入力画面'
    if label.startswith('研究会'):
        return '研究会'
    return label


for r in rows[1:]:
    if not r or r[ct] is None:
        continue
    n += 1
    lab = str(r[ct]).strip()
    t, c_, m = grp(lab), str(r[cc] or '(空)').strip(), str(r[cm] or '(空)').strip()
    by_m_t[(m, t)] += 1; by_m_c[(m, c_)] += 1
    fun_t[t]['CV'] += 1; fun_c[c_]['CV'] += 1; fun_l[lab]['CV'] += 1
    for k, i in flags.items():
        v = r[i]
        if v not in (None, '', 0, '×', '-', 'ー'):
            fun_t[t][k] += 1; fun_c[c_][k] += 1; fun_l[lab][k] += 1
mt = sorted({m for m, _ in by_m_t}); tt = sorted({t for _, t in by_m_t})
sheet('07_実績CV_月×タイプ', ['CV年月'] + tt + ['合計'], [[m] + [by_m_t.get((m, t), 0) for t in tt] + [sum(by_m_t.get((m, t), 0) for t in tt)] for m in mt],
      [10] + [16] * (len(tt) + 1),
      f'出どころ：マーケティング本部｜{BIZ}｜実績.xlsm（流入チャネル別顧客一覧・全{n}件）。メールアドレス・会社名・氏名・電話は読まず、件数だけ集計した。HubSpotのレポートの代わりに、この実績ファイル（CVリスト）を使う。')
cl = sorted({c_ for _, c_ in by_m_c})
sheet('07b_実績CV_月×流入チャネル', ['CV年月'] + cl + ['合計'], [[m] + [by_m_c.get((m, c_), 0) for c_ in cl] + [sum(by_m_c.get((m, c_), 0) for c_ in cl)] for m in mt],
      [10] + [12] * (len(cl) + 1), '流入チャネル別のCV件数（月別）。')
fk = ['CV'] + list(flags)
sheet('07c_実績_商談への進み方', ['区分', '値'] + fk + ['SAL率', 'SQL1率'],
      [['CVタイプ', t] + [fun_t[t][k] for k in fk] + [pct(fun_t[t]['SAL'], fun_t[t]['CV']), pct(fun_t[t]['SQL1'], fun_t[t]['CV'])] for t in tt] +
      [['流入チャネル', c_] + [fun_c[c_][k] for k in fk] + [pct(fun_c[c_]['SAL'], fun_c[c_]['CV']), pct(fun_c[c_]['SQL1'], fun_c[c_]['CV'])] for c_ in cl],
      [12, 18] + [9] * (len(fk) + 2), 'CVタイプ・流入チャネル別の、SAL／SQL／提案／受注／辞退／失注の件数（記入のあるものを数えた）と、CVに対するSAL率・SQL1率。')

fk2 = ['CV'] + list(flags)
sheet('07d_実績_CVの元（メルマガCTA別）', ['CVの元（メルマガ番号・CTA名・資料名）'] + fk2, [[lab] + [fun_l[lab][k] for k in fk2] for lab, _ in sorted(fun_l.items(), key=lambda x: -x[1]['CV'])],
      [64] + [8] * len(fk2), 'CVした元のラベル別の件数（メルマガ番号・CTAの位置・資料名）。どのメルマガ・資料がCV・SALにつながったかの材料。')

# ---------------------------------------------------------------- 既存施策（重複チェック用の件数）
wbs = openpyxl.load_workbook(cp(SHISAKU, 'shisaku.xlsx'), read_only=True, data_only=True)
ex = collections.Counter(); ex_t = collections.Counter(); tot = 0; skipped = 0
for r in list(wbs[C['sheet']].iter_rows(min_row=3, values_only=True)):
    if not r or r[0] in (None, ''):
        continue
    if not r[5] and not r[14]:
        skipped += 1
        continue  # 施策種類もステータスも空の行（No・担当者・URLだけの行）は施策として数えない（アプリ取込と同じ）
    tot += 1
    ex[(str(r[5] or '(空)'), str(r[14] or '(空)'))] += 1
    ex_t[str(r[4] or '(空)')] += 1
sheet('08_既存施策の件数', ['施策種類', 'ステータス', '件数'], [[a, b, v] for (a, b), v in sorted(ex.items())], [18, 12, 8],
      f'出どころ：施策管理表（正本）の{BIZ}シート・施策{tot}行（種類・ステータスが空の{skipped}行は除く）。重複チェックには、施策内容の突き合わせが別途必要（施策内容には事例企業名が含まれるため、この表では件数のみ）。')

# ---------------------------------------------------------------- 閾値（DSP）
disp = {ym: ch[ym].get('Display', 0) for ym in sorted(ch)}
lp = collections.defaultdict(lambda: [0, 0])
for r in lp_rows:
    lp[r[0]][0] += r[3]; lp[r[0]][1] += r[6]
th_rows = []
for ym in sorted(lp):
    th_rows.append([ym, disp.get(ym), lp[ym][0], round(lp[ym][1] / lp[ym][0], 1) if lp[ym][0] else None])
sheet('09_DSP判定の材料', ['年月', 'Displayのセッション', 'DSP流入のLPセッション', 'DSP流入の1セッションあたり滞在(秒)'], th_rows, [10, 18, 20, 26],
      'DSP広告の目的＝認知アップ・エンゲージメントアップ。中島さんの考える判定項目：①セッション数 ②LPの直帰率の低下 ③LPの滞在時間の延長。'
      '【注意】②の直帰率は、現在のダッシュボード／GA4のデータに入っていない（取得項目の追加が必要）。当面は「次ページ遷移率」（月次レポート記載：約15%）とエンゲージメント時間で代替する。'
      '閾値の数値は未決（この表の推移を見て決める）。')

# ---------------------------------------------------------------- 00 概要（先頭に置く）
ov = wb.create_sheet('00_概要', 0)
lines = [
    f'AI分析用 集計表（{BIZ}）  作成日 {TODAY}',
    '',
    '目的：AI（Claude）に渡す材料の確認と、渡す元データそのもの。個人情報（氏名・メール・会社名・電話・受注／SAL／SQL企業名）は含めていない。',
    '範囲：まず1事業（IS）。DSP広告はSCHとISの2事業のみ。スペースクリエイト（SC）の候補は別途（PDCAの施策候補を出すタブに出す）。',
    '',
    'シート一覧',
    '  01 ROI月次：セッション・MQL・SAL・SQL・受注と目標（2026-10まで）',
    '  02 月次レポート所見：10月号のIS（人が書いた所見・示唆を含む。社名は除外）',
    '  03 GEO設問別：ISの22問の質問・回答全文・出現有無・出た企業名。03b 推移、04 定型文（現状のまま）',
    '  05 DSP広告：配信実績（月次・PC/SP別・素材別）。最新ダッシュボード（2026-11-13更新）より',
    '  06 サイトGA4/GSC：チャネル別・DSP流入のLP別・ページ別・検索クエリ',
    '  07 実績CV：実績.xlsm（CVリスト）を件数だけ集計。CVタイプ・流入チャネル別の商談への進み方',
    '  08 既存施策の件数：重複を避けるための現状',
    '  09 DSP判定の材料：セッション・滞在時間（直帰率はデータなし）',
    '',
    '根拠の扱い（すべて「弱い」で出す：中島さん確認済み）',
    '  ・ROIと月次レポートの数字は同じ出どころ（ROIダッシュボード）。根拠は1つと数える。',
    '  ・GEOは1モデル・質問が少ないため「弱い：件数が少ない」。',
    '  ・各候補の根拠に、シート名・期間・数字を必ず書く。',
    '',
    'AIの役割の切り分け（提案。確認待ち）：A 整理（材料にある所見・数字をそのまま候補に）／B 分析（数字から読み取った示唆）／C 新規発案（材料にない新しい施策案）。',
    '出所の表記：A＝月次レポート／ROI／GEO／DSP、B＝AI分析、C＝AI発案。候補の根拠の強さはすべて「弱い」。',
]
for i, t in enumerate(lines, 1):
    c = ov.cell(i, 1, t)
    if i == 1:
        c.font = Font(bold=True, size=14)
ov.column_dimensions['A'].width = 150

os.makedirs(OUTDIR, exist_ok=True)
out = os.path.join(OUTDIR, f'{BIZ}_AI分析用集計表_{TODAY}.xlsx')
wb.save(out)
shutil.rmtree(tmp, ignore_errors=True)
print('saved', out)
for ws in wb.worksheets:
    print(f'{ws.title}: {ws.max_row}行')
