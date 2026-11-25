# -*- coding: utf-8 -*-
r"""AI分析の結果（候補の一覧）をWordにまとめる。候補ファイル（data/candidates/candidates_*.json）から作るので、画面の内容と一致する。
使い方：C:\work\anaconda_install\python.exe -I ai_analysis\make_result_docx.py docs\出力.docx
標準ライブラリだけで .docx を組み立てる（この PC には node・pandoc・LibreOffice・python-docx が無い）。
"""
import sys, os, json, glob, zipfile, re
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = 'Meiryo'
TITLE_TEXT = 'AI分析の結果（全事業）確認資料'


def rpr(size=21, bold=False, color=None):
    s = f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
    if bold:
        s += '<w:b/>'
    if color:
        s += f'<w:color w:val="{color}"/>'
    return s + f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'


def runs(text, size=21, color=None, bold=False):
    out = ''
    for part in re.split(r'(\*\*[^*]+\*\*)', text):
        if not part:
            continue
        b = bold
        if part.startswith('**'):
            part, b = part[2:-2], True
        out += f'<w:r>{rpr(size, b, color)}<w:t xml:space="preserve">{escape(part)}</w:t></w:r>'
    return out


def P(text, size=21, color=None, after=100, bold=False):
    return f'<w:p><w:pPr><w:spacing w:after="{after}"/></w:pPr>{runs(text, size, color, bold)}</w:p>'


def H(level, text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r>{rpr(30 if level == 1 else 24, True, "1F3A5F" if level == 1 else "2F5597")}<w:t>{escape(text)}</w:t></w:r></w:p>'


def TITLE(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r>{rpr(40, True, "1F3A5F")}<w:t>{escape(text)}</w:t></w:r></w:p>'


def B(text):
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/></w:pPr>{runs(text, 21)}</w:p>'


def cell(text, w, head=False, fill=None, size=19):
    shd = ''
    if head:
        shd = '<w:shd w:val="clear" w:color="auto" w:fill="1F3A5F"/>'
    elif fill:
        shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    paras = ''.join(f'<w:p><w:pPr><w:spacing w:after="30"/></w:pPr>{runs(l, size, "FFFFFF" if head else None, head)}</w:p>' for l in str(text).split('\n'))
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}'
            f'<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>{paras}</w:tc>')


def table(widths, head, rows, first_fill=True, size=19, fills=None):
    tw = sum(widths)
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{tw}" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
         '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>')
    x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ''.join(cell(h, widths[i], head=True, size=size) for i, h in enumerate(head)) + '</w:tr>'
    for ri, r in enumerate(rows):
        f0 = (fills[ri] if fills else ('E8EEF7' if first_fill else None))
        x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(cell(c, widths[i], fill=(f0 if i == 0 else None), size=size) for i, c in enumerate(r)) + '</w:tr>'
    return x + '</w:tbl>' + '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>'


# ------------------------------------------------------------------ 候補ファイルを読む
cands = []
for f in sorted(glob.glob(os.path.join(os.path.dirname(HERE), 'data', 'candidates', 'candidates_*.json'))):
    j = json.load(open(f, encoding='utf-8'))
    for c in j['candidates']:
        cands.append(c)
by = {}
for c in cands:
    by.setdefault(c['biz'], []).append(c)
CLS_FILL = {'A': 'E2F0D9', 'B': 'DDEBF7', 'C': 'FFF2CC'}


def cls_letter(c):
    return c['cls'][0]


def cnt(biz, k):
    return sum(1 for c in by.get(biz, []) if cls_letter(c) == k)


ORDER = ['IS', 'SCH', 'MED', 'UN', 'SC']
NAMES = {'IS': 'IS（インサイトセールス）', 'SCH': 'SCH（スクール・学生服）', 'MED': 'MED（メディカルウェア）', 'UN': 'UN（企業ユニフォーム）', 'SC': 'SC（スペースクリエイト）'}
total = sum(len(v) for v in by.values())
body = []
a = body.append
a(TITLE(TITLE_TEXT))
a(P('作成：2026-11-25 ／ 状態：AI分析を全事業（IS・SCH・MED・UN・SC）で実行し、候補ファイルを作成。採否は未実施（中島さんの確認待ち） ／ 関連：確認資料 rev.2 の9章「次の手順」', color='555555'))

a(H(1, '1. 結論（先に）'))
for t in [
    f'**5事業で合計 {total} 件の候補を作りました**（IS {len(by["IS"])}・SCH {len(by["SCH"])}・MED {len(by["MED"])}・UN {len(by["UN"])}・SC {len(by["SC"])}）。すべて状態は「候補」（採否待ち）で、根拠の強さはすべて「弱い」です（確認済み）。採否は行っていません。',
    'Cの新規発案（材料に直接の裏付けがない仮説）は、どの事業も5件の上限内です（' + '・'.join(f'{k} {cnt(k, "C")}' for k in ORDER) + '）。',
    '候補はアプリを開くと**自動で読み込まれ**、施策管理タブの上部に「⏳ 採否待ちの候補：UN 21件／SCH 22件／…」と出ます。事業名をクリックすると、その事業の「採否待ち」が開きます。GitHub・ローカル・Box は同じ内容です。',
    'DSP広告（SCH・IS）の候補には、GA4から取った**直帰率・平均エンゲージメント時間（2026年4月以降）**を書き足しました。ただし**2026年4月に全サイトで計測が変わっている**ため、3月以前とは比べられません（6章）。',
    '月次レポートの所見（区分A）は、**2026年10月号**（FLUED｜2026｜月次定例レポート）から作りました。',
]:
    a(B(t))

a(H(1, '2. 件数'))
rows_ = []
for k in ORDER:
    rows_.append([NAMES[k], str(cnt(k, 'A')), str(cnt(k, 'B')), str(cnt(k, 'C')), str(len(by[k]))])
rows_.append(['合計', str(sum(cnt(k, 'A') for k in ORDER)), str(sum(cnt(k, 'B') for k in ORDER)), str(sum(cnt(k, 'C') for k in ORDER)), str(total)])
a(table([3000, 1800, 1800, 1800, 1800], ['事業', 'A 整理', 'B 分析', 'C 新規発案', '合計'], rows_))

a(H(1, '3. 分析で分かったこと（主なもの）'))
a(H(2, 'IS'))
for t in [
    '**SQL÷SALの32%は、時間のずれを含みます。** 9月はSAL19・SQL0、10月はSAL0・SQL3で、10月のSQLの一部は研究会関連のSALからの引き上がりです（B01）。',
    '**DSP広告はPCとSPで効率が大きく違います**（2026/10：クリック単価 PC約177円・SP約19円）。目的が認知でもあるため、振り替えは量と期間を区切って試す案にしました（B05）。',
    '**メルマガのCTAは、WPとフッターのSAL率が高い**（WP系25.0%・フッター30.8%・本文cta1 12.3%）。ABテスト案にしました（B03）。',
    '**GEOの「出現」は、認知の根拠になりません。** 企業情報直接型は質問に社名が入っているためで、回答は「確実な情報を持ち合わせていない」と述べています（G07）。',
]:
    a(B(t))
a(H(2, 'SCH'))
for t in [
    '**DSPは、クリックが急増したのに、LPの滞在が短い**（クリック単価は6月105円→9月8.3円、クリックは1,915→45,617（9月）。LPの平均エンゲージメント時間は2.7〜4.6秒）。9月は展示会予約ページが77.3%を占め、直帰率はページの種類に左右されます（B01）。',
    '**展示会キャンペーン（東京・大阪）は、通常キャンペーンよりクリック率が高い**（9月：CTR 1.30%・クリック単価7.1円／通常 0.85%・10.3円）。広告グループ名と月次レポートの数字からの推定です（B02）。',
    '**MQLの質**：10月のMQL6件のうち5件は、ECサイトにアクセスできないという問い合わせでした（月次レポート p.17）。学校向けのMQLは1件の可能性があります（B06）。',
    'サイト流入は月間100程度で、指名検索の表示数・流入数が少なく、事業認知度が低いことが課題です（月次レポート p.25）。プッシュ型で接触を増やせているのは、DSP広告とメルマガだけです。',
]:
    a(B(t))
a(H(2, 'MED'))
for t in [
    '**流入が減っています**：3〜10月のセッションは前年同期比 −16.6%（19,526／23,405）。GA4の自然検索は10月が前年同月比 −39.6%で、掲載順位はほぼ変わっていません（B01）。',
    '**カタログ請求フォームのCVが約45%減りました**（2026-01〜04の月平均20.5件→06〜10の月平均11.2件。セッションは−10.9%）。フォームの項目や導線の確認が必要です（B03）。',
    '**メルマガ：要点を絞ったショート版の事例紹介が、クリック率7.3%で勝ちました**（No.68のABテスト。月次レポート p.34）。標準化を候補にしました（A04）。',
    '**GEO**：自社が出たのは32問中2問だけで、いずれも質問に社名が入る問いです。ブランド名の表記が揺れています（G07）。',
]:
    a(B(t))
a(H(2, 'UN'))
for t in [
    '**セッションは前年同期比 +26.7%、MQLは −11.2%**：MQL÷セッションは0.82%から0.57%に下がりました（月次目標25件に届いたのは5月・7月・10月の3か月のみ）。',
    '**メルマガの誘導先別SAL率**：事前準備チェックシート（WP2）系が20件中6件（30.0%）で最も高く、WP3系は10件でSAL 0でした（B02）。',
    '**リスティング経由のSAL率は9.8%**（オーガニック検索27.1%、全体20.2%）。ただし10月は「ユニフォーム デザイン」経由の問い合わせが好調と、月次レポートにあります（B03・A06）。',
    '**事例ページが閲覧の73%を占める**一方、滞在は短く、資料DL・相談への導線がありません（B04）。',
    '受注2件は、CVから商談化まで半年以上かかった例を含みます（月次レポート p.6）。',
]:
    a(B(t))
a(H(2, 'SC'))
for t in [
    '**SCは、ROI・GA4・GSC・実績ファイル・月次レポート・施策管理表のデータが、そもそも存在しません**（サイトが制作中で未公開。ROIは4事業のみ）。集計表は、GEOの10問と、サイト制作の下書き（2026-07-13更新。写真は未配置）から作りました。',
    '**GEO**：自社は15回の実行すべてで0/10でした。**サイト下書き**には、強み（ユニフォームまでトータル／サステナブルな空間演出など）・進め方4ステップ・予定のコラム3本・資料2種がありますが、**事例の枠は写真・中身が未配置**です（B05）。',
]:
    a(B(t))

a(H(1, '4. 候補の一覧'))
a(P('区分の色：A 整理＝緑／B 分析＝青／C 新規発案＝黄。根拠には、シート名・期間・数字を書いています（略称：01〜10＝集計表のシート番号）。推奨担当は、過去の同じ施策種類の担当チームから付けた案で、個人名は未定です（未定の候補は「未定」）。', color='555555'))
W = [900, 700, 1100, 4300, 5600, 1900]
HEAD = ['起票元ID', '区分', '出所', '施策内容（種類／目的）', '根拠', '推奨担当・検証日・検証指標']
for i, biz in enumerate(ORDER, 1):
    a(H(2, f'4-{i}. {NAMES[biz]}'))
    rows, fills = [], []
    for c in by[biz]:
        idshort = c['originId'].replace('AI-%s-20261125-' % biz, '')
        parent = f'\n（次サイクル：親施策No.{c["parentNo"]}）' if c.get('parentNo') else ''
        rows.append([
            idshort, c['cls'], c['proposedFrom'],
            f'**{c["cat"]}｜{c["mokuteki"]}**\n{c["naiyo"]}{parent}',
            c['evidence'],
            f'担当：{c.get("tanto") or "未定"}\n検証日：{c["check"]}\n指標：{c["verifyKind"]}',
        ])
        fills.append(CLS_FILL[cls_letter(c)])
    a(table(W, HEAD, rows, size=17, fills=fills))

a(H(1, '5. GEOの22問（IS）と候補の対応'))
a(P('ISは、すべての設問をいずれかの候補に割り当てました（ID末尾の番号は設問IDの末尾）。ほかの事業のGEOの対応は、各候補の根拠に設問IDを書いています（SCH 10問・MED 32問・UN 24問・SC 10問）。', color='555555'))
GEO_MAP = [
    ['G01', '業者推薦型・比較・選定型', '043、048、S2-07、S2-11', '選び方・一気通貫ガイド'],
    ['G02', '業者推薦型（大手・インナー・OEM）', '044、052、S2-08（回答が空）、S2-10、S2-12', '大手企業・インナーブランディング・OEMの紹介ページ'],
    ['G03', '事例・実績型（既存の4事例）', 'S2-03、S2-04、S2-05、S2-06', '既存事例ページ（ANA・ファンケル・ブリヂストン・カープ）の見直し'],
    ['G04', '事例・実績型・課題解決型（アップサイクル／インナー）', '050、051、S2-09', '事例まとめ（ハブ）記事2本'],
    ['G05', '課題解決型', '045、046、047', '課題起点のコラム3本'],
    ['G06', 'トレンド・市場情報型', '049', 'トレンド記事（年1回更新）'],
    ['G07', '企業情報直接型', 'S2-01、S2-02', '公式な紹介情報の整備'],
]
a(table([1000, 3600, 4600, 5300], ['候補', '質問タイプ', '設問（末尾）', '施策案'], GEO_MAP))

a(H(1, '6. 注意点・限界（データの確認が必要なこと）'))
a(H(2, '全事業に共通'))
for t in [
    '**2026年4月に、全サイト・全チャネルで直帰率が一斉に下がり、閲覧ページ数/セッションがほぼ倍になっています**（例：ISのサイト全体は、直帰率56%→26%、閲覧ページ数2.58→4.13。UN・SCH・MEDも同様。DSP以外のチャネルも同時）。広告の効果ではなく、GTMタグやGA4の計測の変更（ページ表示の二重計測など）の可能性が高いです。**DSPの直帰率は4月以降の月だけで比べてください。** 4月前後のGTM・GA4の変更履歴の確認が必要です。',
    '**GEOは1つのAIモデルの結果で、質問は事業ごとに10〜32件**です。回答に出た社名の実在性・正確性は未確認です（根拠はすべて「弱い」）。',
    '**既存施策との重複は、件数でしか確認できていません。** 候補ごとに、採否の前に施策管理表で重複を確認してください（施策内容に事例企業名を含むため、AIには渡していません）。次サイクルの候補（IS：No.178、SCH：No.136・146、MED：No.149・151、UN：No.175）は、親施策の現在の状態の確認が必要です。',
    '**実績ファイルの流入チャネルの表記に揺れがあります**（organic／オーガニック検索、mail／Eメールなど。`#REF!` が混ざる事業もあり）。**CVの種類の分け方も、2025年6月と7月の間で変わっています**（種類別の月次比較は2025年7月以降のみ有効）。',
    'SAL・SQL・提案は、実績ファイルの記入の有無を数えたものです。最近のCVは、まだ進んでいない可能性があります。',
]:
    a(B(t))
a(H(2, '事業別'))
for t in [
    'IS：GEO定型文の「想起された競合」（チトセ・ユニフォームネット）はユニフォーム領域の社名で、ISの回答に出た社名と一致しません。GEO設問 Q-S2-C4-D3-08 は回答が空でした。月次レポートの「指名検索クリック0」と、GSCの指名検索のクリックは、定義が違うように見えます。',
    'SCH：GEO定型文（04）が空で、競合が実態に合っているか確認できませんでした。回答に「オンワード樫山」が3問に出ていますが、全問「出なかった」判定です（判定が自社名の一致だけの可能性）。展示会キャンペーン＝「東京・大阪」、通常＝「office_SCH_pc/sp」は、グループ名と月次レポートの数字からの推定です。10b（LP別）と10（月合計）のセッションは、集計方法の違いで少しずれます。',
    'MED：実績ファイルの「CV年月」に、月ではなく日時が入っていました。集計スクリプトで月に直しています。メルマガNo.65・68のクリック率は、月次レポートの記載と、送信数×クリック数の再計算が合いません（集計の範囲の違いの可能性）。',
    'UN：10月のメルマガNo.40・41は、月次レポートではCV0ですが、実績ファイルのCVの元ラベルでは10件・1件あります。10月のEmail経由セッションとコラム1位の表示には、BOTのスキャンの疑いがあります（月次レポート p.7の仮説）。GEO定型文に自重堂がありませんが、実際の回答には多く出ています。',
    'SC：ROI・GA4・GSC・実績のデータは存在しません（サイト未公開。集計表の12シート）。根拠はGEOとサイト制作の下書き（公開前の計画で、効果のデータではない）で、SCに該当の案件があるか（B02・B05）は未確認です。サイトを公開してGA4・GSCを取り始めたら、集計表を追加します。',
]:
    a(B(t))

a(H(1, '7. 画面での表示と、しくみ'))
for t in [
    '**候補の自動読込**：アプリを開いたとき、HTMLと同じ場所の `data/candidates/candidates.js`（全候補ファイルをまとめた索引）を読み、状態「候補」の行として各事業のシートに追加します。取り込む操作・画面はありません。フォルダの許可も不要です。起票元IDが同じ候補・以前に読み込んだ候補・非採用にした候補は、追加しません。',
    '**画面の上部の案内**：採否待ちの候補の件数を事業別に出し、クリックでその事業の「採否待ち」を開きます。採否は行を選んで、右の詳細パネルのボタン（理由は必須）で行います。',
    '**SC用の置き場**：上の事業タブの「SC（準備中）」と連動する5つ目のシートです（Excel取込・出力の対象外）。',
    '**候補は通知の対象外**：推奨検証日を入れても、状態が「候補」の間は通知しません。',
    '**候補ファイルの作り方**：`ai_analysis/make_candidates*.py` が `data/candidates/candidates_<事業>_<日付>.json` と索引 `candidates.js` を作ります（git・Box に同じものがあります）。集計表は `build_ai_input.py <事業>` で作ります。最終形は、毎営業日のバッチが同じ候補ファイルを「まとめたデータ」へ追加します（バッチは未実装）。',
]:
    a(B(t))

a(H(1, '8. 確認したいこと・次の手順'))
a(table([6200, 8370], ['確認していただきたいこと', '補足'], [
    [f'候補の内容の確認（5事業・{total}件）と、採否の流れのお試し', 'アプリの「採否待ち」で確認してください。**採否は中島さんが行います（AIは行いません）。**'],
    ['4月前後の GTM・GA4 の変更履歴の確認', '直帰率の変化が、計測の変更によるものかを確かめるためです。確認できるまで、DSPの直帰率は4月以降の月だけで見てください。'],
    ['（継続）GA4のデバイス別（PC／SP）の直帰率を取得する追加', 'PC→SPの配分見直し（IS B05、SCH B02）の効果を、エンゲージメントで見るために必要です。Apps Scriptへの追加になります。'],
]))
a(P('次の手順：中島さんが候補を確認して、採否を行う。採用した候補は状態「未着手」になり、施策管理表と同じシートに載ります。'))

footer_xml = (f'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
              f'<w:r>{rpr(16, False, "777777")}<w:t xml:space="preserve">{TITLE_TEXT}　</w:t></w:r>'
              f'<w:r>{rpr(16, False, "777777")}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{rpr(16, False, "777777")}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
              f'<w:r>{rpr(16, False, "777777")}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{rpr(16, False, "777777")}<w:t>1</w:t></w:r><w:r>{rpr(16, False, "777777")}<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>' + ''.join(body) +
            '<w:sectPr><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            '<w:pgMar w:top="1000" w:right="1134" w:bottom="1000" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')

styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="320" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="220" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5597"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>'''

numbering = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="540" w:hanging="270"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'''
rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'''
doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'''
core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{TITLE_TEXT}</dc:title><dc:creator>KPI-PDCA</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-11-25T00:00:00Z</dcterms:created></cp:coreProperties>'''

out = sys.argv[1]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', document)
    z.writestr('word/styles.xml', styles)
    z.writestr('word/numbering.xml', numbering)
    z.writestr('word/footer1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + footer_xml)
    z.writestr('word/_rels/document.xml.rels', doc_rels)
    z.writestr('docProps/core.xml', core)
print('written', out)
