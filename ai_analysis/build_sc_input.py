# -*- coding: utf-8 -*-
"""SC（スペースクリエイト）のAI分析用集計表を作る（個人情報なし）。

使い方：C:\\work\\anaconda_install\\python.exe -I ai_analysis\\build_sc_input.py
出力：ai_analysis/input/SC_AI分析用集計表_<日付>.xlsx

SCは、ROI・GA4・GSC・実績ファイル・施策管理表のシートが「まだ存在しない」ため（サイトは制作中で未公開）、
使える材料は次の2つ：
  (1) GEO（空間ビジネス領域 C6 の10問）
  (2) サイト制作の下書き（社内ドキュメントの「スペースクリエイト｜サイト制作｜1_service-site.html／2_lp.html」。2026-07-13更新。
      写真は未配置）。ページの文言（訴求・構成・CTA・予定のコラムと資料）を要点として集計表に載せる。
サイトを公開してGA4・GSCを取り始めたら、build_ai_input.py に SC を足して、IS・SCHなどと同じ形にする。
"""
import os, re, csv, json, glob, shutil, tempfile, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

TODAY = datetime.date.today().isoformat()
DOC = r'C:\Users\612316\Box\マーケティンググループ\デジタルマーケティングDiv\マーケティング本部\社内ドキュメント'
GEO = r'C:\Users\612316\Documents\GitHub\GEO\monitoring\data'
HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(HERE, 'input')
DOMAIN = '空間ビジネス（スペースクリエイト）'
SITE_FILES = [('サイト', 'マーケティング本部｜スペースクリエイト｜サイト制作｜1_service-site.html'),
              ('LP', 'マーケティング本部｜スペースクリエイト｜サイト制作｜2_lp.html')]
tmp = tempfile.mkdtemp(prefix='sc_input_')
wb = openpyxl.Workbook(); wb.remove(wb.active)
HDR = PatternFill('solid', fgColor='1F3A5F')


def sheet(title, header, rows, widths=None, note=None):
    ws = wb.create_sheet(title)
    r0 = 1
    if note:
        ws.cell(1, 1, note).font = Font(italic=True, color='555555')
        ws.cell(1, 1).alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(header)))
        ws.row_dimensions[1].height = 60
        r0 = 2
    for j, h in enumerate(header, 1):
        c = ws.cell(r0, j, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for i, row in enumerate(rows, r0 + 1):
        for j, v in enumerate(row, 1):
            ws.cell(i, j, v).alignment = Alignment(wrap_text=isinstance(v, str) and len(v) > 40, vertical='top')
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths[j - 1] if widths and j - 1 < len(widths) else 14)
    ws.freeze_panes = ws.cell(r0 + 1, 1)


def truthy(v):
    return str(v).strip().lower() in ('true', '1', 'yes')


# ------------------------------------------------------------ GEO
ins = json.load(open(os.path.join(GEO, 'reports', 'insights_20261104_133016.json'), encoding='utf-8'))
res = list(csv.DictReader(open(os.path.join(GEO, 'results', 'results_20261104_133016_r1.csv'), encoding='utf-8-sig')))
geo_rows = [[r['question_id'], r['type_label'], r['stakeholder_label'], r['question'], '出た' if truthy(r['mention_detected']) else '出なかった',
             r['mention_position'], r['entities_found'], r['urls_found'], r['specificity_tier'], r['model_id'], r['answer']]
            for r in res if r['domain_label'] == DOMAIN]
sheet('03_GEO_設問別', ['設問ID', '質問タイプ', '想定読者', '質問', '自社の出現', '出現位置', '回答に出た企業名', '回答に出たURL', '具体度', 'モデル', '回答（全文）'],
      geo_rows, [10, 14, 12, 50, 10, 8, 40, 30, 8, 14, 120],
      f'出どころ：GEO insights/results {ins["generated_at"]}（2026-11-04実行）。{DOMAIN}の全{len(geo_rows)}問。全問が1モデル（claude-sonnet-5）の結果。'
      '競合の社名はAIに渡してよい（中島さん確認済み）。質問の件数が少なく1モデルのため、GEO由来の候補は根拠を「弱い」とする。「回答に出た企業名」列は空のことが多いので、回答本文から読み取る。')
hist = []
types = sorted({r['type_label'] for r in res if r['domain_label'] == DOMAIN})
for f in sorted(glob.glob(os.path.join(GEO, 'results', 'results_*_r1.csv'))):
    rows = list(csv.DictReader(open(f, encoding='utf-8-sig')))
    tot = collections.Counter(); det = collections.Counter()
    for r in rows:
        if r['domain_label'] != DOMAIN:
            continue
        tot[r['type_label']] += 1
        if truthy(r['mention_detected']):
            det[r['type_label']] += 1
    if not sum(tot.values()):
        continue
    p = os.path.basename(f).split('_')
    hist.append([p[1] + '_' + p[2], sum(det.values()), sum(tot.values())] + [f'{det[t]}/{tot[t]}' for t in types])
sheet('03b_GEO_推移', ['実行日', '出現', '質問数'] + types, hist, [16, 7, 8] + [14] * len(types), f'{DOMAIN}の出現の推移（質問タイプ別の「出現／質問数」）。')
pb = [[a['type'], f'{a["miss"]}/{a["total"]}', a['priority'], a['playbook'], '、'.join(a.get('cover_attributes', [])), '、'.join(a.get('beating_competitors', []))]
      for a in ins['actions'] if a['domain'] == DOMAIN]
sheet('04_GEO_定型文', ['質問タイプ', '出なかった／質問数', '優先度', '定型文（現状のまま）', '回答で重視された観点', '想起された競合'], pb, [14, 14, 8, 80, 50, 40],
      '定型文はいったん現状の表現のまま。SCは、GEOのinsights（改善アクション）に領域が無いため、定型文は空。候補の内容は、上の「設問別」の結果からの分析にする。')

# ------------------------------------------------------------ サイト制作の下書き（ページの文言を、出てくる順に）
site_rows = []
for key, fn in SITE_FILES:
    p = os.path.join(DOC, fn)
    cp = os.path.join(tmp, key + '.html'); shutil.copy2(p, cp)
    t = open(cp, encoding='utf-8', errors='replace').read()
    m = re.search(r'<script type="__bundler/template">(.*?)</script>', t, flags=re.S)
    tpl = json.loads(m.group(1).strip())
    runs = re.findall(r'[\u3040-\u30ff\u4e00-\u9fff][\u3040-\u30ff\u4e00-\u9fffA-Za-z0-9・、。！？「」『』（）()／\-ー〜～%％,.:：\s]{3,}', tpl)
    seen = []
    for r in runs:
        r = re.sub(r'\s+', ' ', r).strip()
        if len(r) >= 5 and r not in seen:
            seen.append(r)
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d')
    for i, r in enumerate(seen, 1):
        placeholder = 'ドラッグ' in r
        site_rows.append([key, mt, i, r, '写真の置き場（未配置）' if placeholder else ''])
sheet('11_サイト制作の下書き', ['資料', '更新日', '出現順', 'ページの文言（要約せず原文のまま）', '備考'], site_rows, [8, 11, 8, 110, 22],
      '出どころ：社内ドキュメント「マーケティング本部｜スペースクリエイト｜サイト制作｜1_service-site.html／2_lp.html」（2026-07-13更新。いずれも制作中の下書きで、公開前）。'
      'ページの文言を出てくる順に抜き出した。「…の写真をドラッグ」は、写真が未配置の枠。事例・実績の中身（企業名・成果・点数）は、まだ入っていない。会社名・住所など公開情報のみで、個人情報は含まない。')

# ------------------------------------------------------------ データの有無
sheet('12_データの有無', ['データ', 'ある／ない', '理由・補足'], [
    ['ROI（セッション・MQL・SAL・SQL・受注・コスト）', 'ない', 'ROIダッシュボードの事業は UN・SCH・MED・IS の4つ。SCは対象外'],
    ['GA4（サイトのアクセス）', 'ない', 'SCのサイトは制作中で未公開。DSPダッシュボード・Apps Scriptに、SCのGA4プロパティが無い'],
    ['GSC（検索）', 'ない', '同上（サイト未公開）'],
    ['実績ファイル（CVリスト）', 'ない', '「マーケティング本部｜SC｜実績.xlsm」は存在しない'],
    ['月次レポート（FLUED）', 'ない', '2026年10月号は UN・SCH・MED・IS の4事業のみ'],
    ['施策管理表（既存施策）', 'ない', 'シートは UN・SCH・MED・IS の4つ。SCの既存施策は無い（アプリの「SC」シートは、候補の置き場）'],
    ['DSP広告', 'ない', 'DSPはSCHとISのみ'],
    ['GEO（空間ビジネス領域 C6）', 'ある', '10問・1モデル。2026-07-16以降の実行の推移あり（03・03b）'],
    ['サイト制作の下書き', 'ある', '11シート。2026-07-13更新。公開前で、事例の中身はまだ入っていない'],
], [44, 12, 80], 'SCの分析に使える材料の一覧。ない材料は、サイトの公開後に追加する（build_ai_input.py に SC を足す）。')

ov = wb.create_sheet('00_概要', 0)
lines = [
    f'AI分析用 集計表（SC＝スペースクリエイト）  作成日 {TODAY}',
    '',
    '目的：AI（Claude）に渡す材料の確認と、渡す元データそのもの。個人情報は含まない。',
    '範囲：SCは、サイト制作中で未公開のため、ROI・GA4・GSC・実績ファイル・月次レポート・施策管理表のデータが無い（12シート）。材料は、GEOとサイト制作の下書きだけ。',
    '',
    'シート一覧',
    '  03 GEO設問別：空間ビジネス領域の10問の質問・回答全文・出現有無。03b 推移、04 定型文（SCは空）',
    '  11 サイト制作の下書き：サービスサイト・LPのページの文言（訴求・構成・CTA・予定のコラムと資料）。事例の中身は未入力',
    '  12 データの有無：ない材料と、その理由',
    '',
    '根拠の扱い（すべて「弱い」で出す：中島さん確認済み）',
    '  ・GEOは1モデル・質問が少ないため「弱い：件数が少ない」。',
    '  ・サイト制作の下書きは、公開前の計画で、効果のデータではない（計画との差を指摘する根拠には使えるが、効果の根拠にはならない）。',
    '  ・各候補の根拠に、シート名・数字・出どころを必ず書く。',
    '',
    'AIの役割の区分：A 整理（材料にある所見・数字をそのまま候補に）／B 分析（読み取った示唆）／C 新規発案（材料にない新しい施策案。1事業5件まで）。',
]
for i, t in enumerate(lines, 1):
    c = ov.cell(i, 1, t)
    if i == 1:
        c.font = Font(bold=True, size=14)
ov.column_dimensions['A'].width = 150
os.makedirs(OUTDIR, exist_ok=True)
out = os.path.join(OUTDIR, f'SC_AI分析用集計表_{TODAY}.xlsx')
wb.save(out)
shutil.rmtree(tmp, ignore_errors=True)
print('saved', out)
for ws in wb.worksheets:
    print(f'{ws.title}: {ws.max_row}行')
