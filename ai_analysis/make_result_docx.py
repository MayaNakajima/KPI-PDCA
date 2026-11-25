# -*- coding: utf-8 -*-
r"""AI分析の結果（候補の一覧）をWordにまとめる。候補ファイル（data/candidates/candidates_*.json）から作るので、画面の内容と一致する。
使い方：C:\work\anaconda_install\python.exe -I ai_analysis\make_result_docx.py docs\出力.docx
標準ライブラリだけで .docx を組み立てる（この PC には node・pandoc・LibreOffice・python-docx が無い）。
"""
import sys, os, json, glob, zipfile, re
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = 'Meiryo'
TITLE_TEXT = 'AI分析の結果（IS・SC）確認資料'


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


body = []
a = body.append
a(TITLE(TITLE_TEXT))
a(P('作成：2026-11-25 ／ 状態：AI分析（IS・SC）を実行し、候補ファイルを作成。採否は未実施（中島さんの確認待ち） ／ 関連：確認資料 rev.2 の9章「次の手順」2・3', color='555555'))

a(H(1, '1. 結論（先に）'))
for t in [
    f'**IS 22件・SC 4件、合計 {len(cands)} 件の候補を作りました。** すべて状態は「候補」（採否待ち）です。根拠の強さは、すべて「弱い」です（確認済み）。',
    f'区分の内訳：IS＝A 整理 {cnt("IS","A")}件／B 分析 {cnt("IS","B")}件／C 新規発案 {cnt("IS","C")}件、SC＝B 分析 {cnt("SC","B")}件／C 新規発案 {cnt("SC","C")}件。Cは1事業5件までの上限内です。',
    'GEOの22問（IS）は、すべていずれかの候補に割り当てました（5章）。SCは、材料がGEOの空間ビジネス領域の10問しか無いため、根拠はGEOのみです。',
    '候補は、アプリを開くと「保存先フォルダの data/candidates/」から**自動で読み込まれ**、IS・SCのシートの「⏳ 採否待ち」に出ます（取り込む操作・画面はありません）。SCは上の事業タブの「SC」と連動する5つ目のシートです。',
    '**まだ Box の運用フォルダには何も置いていません。** 置いてよいかのご確認をお願いします（8章）。',
]:
    a(B(t))

a(H(1, '2. 件数'))
a(table([2400, 1800, 1800, 1800, 1800], ['事業', 'A 整理', 'B 分析', 'C 新規発案', '合計'], [
    ['IS（インサイトセールス）', str(cnt('IS', 'A')), str(cnt('IS', 'B')), str(cnt('IS', 'C')), str(len(by['IS']))],
    ['SC（スペースクリエイト）', str(cnt('SC', 'A')), str(cnt('SC', 'B')), str(cnt('SC', 'C')), str(len(by['SC']))],
]))

a(H(1, '3. 分析で気づいたこと（候補の前提）'))
for t in [
    '**SQL÷SALの32%は、時間のずれを含みます。** 9月はSAL19・SQL0、10月はSAL0・SQL3で、10月のSQL3件のうち2件は外部団体の研究会関連のSALからの引き上がりです（月次レポートp.41）。9月のSAL19件の追跡が、最初の一手になります（候補 B01）。',
    '**DSP広告は、PCとSPで効率が大きく違います。** 2026/10は、PCが費用の69.9%を使ってクリックの19.8%、クリック単価は約177円。SPは約19円です（05b）。ただし目的が認知（imp）でもあるため、予算の振り替えは量と期間を区切って試す案にしました（B05）。',
    '**メルマガのCTAは、本文冒頭のサービス資料より、WPとフッターのSAL率が高い**（WP系25.0%・フッター30.8%・本文cta1 12.3%。07d）。件数・配信号の差を含むため、ABテスト案にしました（B03）。',
    '**GEOの「出現」は、認知されている根拠になりません。** 企業情報直接型は「出現」2/2ですが、質問に社名が入っているためで、回答は「確実な情報を持ち合わせていない」と述べています（G07）。',
    '**DSPは2026/07から広告グループ名が変わり、クリック数は増えたのに、DSP経由のCVは全月0、MQL÷セッションは9月1.04%→10月0.15%です**（05・01）。狙う企業を絞る案（C03）は、材料に直接の裏付けが無いため、区分Cの仮説にしています。',
]:
    a(B(t))

a(H(1, '4. 候補の一覧'))
a(P('区分の色：A 整理＝緑／B 分析＝青／C 新規発案＝黄。根拠には、シート名・期間・数字を書いています（略称：01〜09＝集計表のシート番号、07d＝実績ファイルの集計）。推奨担当は、過去の同じ施策種類の担当チームから付けた案で、個人名は未定です（SCは未定）。', color='555555'))
W = [900, 700, 1100, 4300, 5600, 1900]
HEAD = ['起票元ID', '区分', '出所', '施策内容（種類／目的）', '根拠', '推奨担当・検証日・検証指標']
for biz, title in [('IS', '4-1. IS（インサイトセールス）'), ('SC', '4-2. SC（スペースクリエイト）')]:
    a(H(2, title))
    rows, fills = [], []
    for c in by[biz]:
        idshort = c['originId'].replace('AI-%s-20261125-' % biz, '')
        parent = f'\n（次サイクル：親施策No.{c["parentNo"]}）' if c.get('parentNo') else ''
        rows.append([
            idshort,
            c['cls'],
            c['proposedFrom'],
            f'**{c["cat"]}｜{c["mokuteki"]}**\n{c["naiyo"]}{parent}',
            c['evidence'],
            f'担当：{c.get("tanto") or "未定"}\n検証日：{c["check"]}\n指標：{c["verifyKind"]}',
        ])
        fills.append(CLS_FILL[cls_letter(c)])
    a(table(W, HEAD, rows, size=17, fills=fills))

a(H(1, '5. GEOの22問（IS）と候補の対応'))
a(P('すべての設問をいずれかの候補に割り当てました。ID末尾の番号は、GEOの設問ID（Q-C4-…／Q-S2-…）の末尾です。', color='555555'))
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
for t in [
    '**GEOは1つのAIモデルの結果で、質問は1〜8件**です。回答に出た社名の実在性・正確性は未確認です（根拠はすべて「弱い」）。',
    '**GEOの定型文の「想起された競合」（チトセ・ユニフォームネットなど）は、ユニフォーム領域の社名**で、ISの回答に実際に出た社名（広告代理店・印刷会社・ノベルティ専門会社）とは一致しません。定型文の元のデータの確認が必要です。',
    '**GEO設問 Q-S2-C4-D3-08 は回答が空**でした。再取得が必要です（G02に含めていますが、分析は他の4問によります）。',
    '**月次レポートの「指名検索クリック0（9月）」と、GSCの上位クエリ（指名検索に71クリックなど）は、定義が違う**ように見えます。リスティングの候補（A06）の前提として、確認が必要です。',
    '**設問と事例ページの対応（G03）は、名称の一致からの推定**です。集計表に対応表がありません。',
    '**既存施策との重複は、件数（08）でしか確認できていません。** 候補ごとに、採否の前に施策管理表で重複を確認してください（施策内容に事例企業名を含むため、AIには渡していません）。',
    '**SCは、ROI・GA4・GSC・実績のデータを集計していません。** 根拠はGEOの10問のみで、SCにその案件があるか（B02）も未確認です。',
    '**LPの直帰率は、いまのデータに無い**ため、DSPの候補（B04・A04）の検証指標はエンゲージメント時間と次ページ遷移率にしています（直帰率の追加は、別途ご判断をお願いしています）。',
]:
    a(B(t))

a(H(1, '7. 画面での表示と、できたしくみ'))
for t in [
    '**候補の自動読込**：アプリを開いたとき、保存先フォルダ（「🗄バックアップ設定」で選んだフォルダ）の data/candidates/*.json を読み、状態「候補」の行として該当事業のシートに追加します。起票元IDが同じ候補・以前に読み込んだ候補・非採用にした候補は、追加しません。フォルダへのアクセス許可が切れている場合は、最初のクリックのときに1回だけ許可を求めます（許可を求める画面は、ブラウザの標準のものです）。',
    '**SC用の置き場**：上の事業タブにあった「SC（準備中）」と連動する5つ目のシートを追加しました。列はISと同じです。Excelの施策管理表にシートが無いため、Excelの取込・出力の対象外です。',
    '**出所の選択肢**：「月次レポート」「AI発案」を追加しました（区分A＝月次レポート／ROI／GEO／DSP、B＝AI分析、C＝AI発案）。',
    '**候補は通知の対象外**：推奨検証日を入れても、「検証日が近い」の通知には含めません（状態が「候補」の間）。',
    '**最終形**：毎営業日のバッチが、同じ候補ファイルを「まとめたデータ」へ追加します（バッチは未実装。それまでの暫定が上の自動読込です）。',
]:
    a(B(t))

a(H(1, '8. 確認したいこと・次の手順'))
a(table([6200, 8370], ['確認していただきたいこと', '補足'], [
    ['候補ファイル（IS・SCの2つ）を Box の `生成AI\\KPI-PDCA\\data\\candidates\\` に置いてよいか', '置くと、アプリを開いたときにシートの「採否待ち」に出ます。Boxの運用フォルダへの反映なので、事前の確認をお願いします。候補ファイルは git の data/candidates/ にあります（Box の運用フォルダとも同じ構成にそろえます）。'],
    ['候補の内容の確認（IS 22件・SC 4件）', '4章の一覧、またはアプリの「採否待ち」で確認。採否（理由つき）の流れもお試しください。**採否は、中島さんが行います（AIは行いません）。**'],
    ['（前回からの確認事項）DSPのLPの直帰率をGA4の取得項目に追加してよいか', 'Apps Script／ダッシュボードの変更を伴います。今回は変更していません。'],
]))
a(P('次の手順：中島さんが候補を確認し、採否の流れを試す。問題なければ、SCH・MED・UN・SC分の集計表を作る（`build_ai_input.py` の BIZ_CONF と月次レポート所見を足す）。'))

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
