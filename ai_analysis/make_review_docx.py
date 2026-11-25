# -*- coding: utf-8 -*-
r"""AI分析の判断材料 確認資料（Word）を作る。標準ライブラリだけで .docx を組み立てる（この PC には node・pandoc・LibreOffice がないため）。
使い方：C:\work\anaconda_install\python.exe -I ai_analysis\make_review_docx.py docs\出力.docx
内容（本文）はこのファイル内の a(...) の並び。直したら再実行して .docx を作り直す。日本語を含むスクリプトは、ヒアドキュメントではなくファイルとして保存して実行する。"""
import sys, zipfile, re
from xml.sax.saxutils import escape

FONT = 'Meiryo'
W = 14570  # 本文幅（DXA）


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


def P(text, size=21, color=None, after=100, bold=False, align=None):
    jc = f'<w:jc w:val="{align}"/>' if align else ''
    return f'<w:p><w:pPr><w:spacing w:after="{after}"/>{jc}</w:pPr>{runs(text, size, color, bold)}</w:p>'


def H(level, text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr><w:r>{rpr(30 if level == 1 else 24, True, "1F3A5F" if level == 1 else "2F5597")}<w:t>{escape(text)}</w:t></w:r></w:p>'


def TITLE(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r>{rpr(40, True, "1F3A5F")}<w:t>{escape(text)}</w:t></w:r></w:p>'


def B(text):
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/></w:pPr>{runs(text, 21)}</w:p>'


def cell(text, w, head=False, fill=None):
    shd = ''
    if head:
        shd = '<w:shd w:val="clear" w:color="auto" w:fill="1F3A5F"/>'
    elif fill:
        shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    paras = ''.join(f'<w:p><w:pPr><w:spacing w:after="30"/></w:pPr>{runs(l, 19, "FFFFFF" if head else None, head)}</w:p>' for l in str(text).split('\n'))
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}'
            f'<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr>{paras}</w:tc>')


def table(widths, head, rows, first_fill=True):
    tw = sum(widths)
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{tw}" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
         '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>')
    x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + ''.join(cell(h, widths[i], head=True) for i, h in enumerate(head)) + '</w:tr>'
    for r in rows:
        x += '<w:tr>' + ''.join(cell(c, widths[i], fill=('E8EEF7' if i == 0 and first_fill else None)) for i, c in enumerate(r)) + '</w:tr>'
    return x + '</w:tbl>' + '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>'


body = []
a = body.append
a(TITLE('AI分析の判断材料 確認資料（rev.2）'))
a(P('作成：2026-11-25 ／ rev.1 に中島さんの確認結果を反映 ／ 状態：確認待ち（AI分析は未実行）／ 関連：共通資料 7章（AI示唆）、仕様書 3-0節（候補・採否）', color='555555'))
a(P('集計表（IS分）：ai_analysis/input/IS_AI分析用集計表_2026-11-25.xlsx（個人情報なし）', color='555555'))

a(H(1, '1. 結論（先に）'))
for t in [
    '**AIの役割は「整理」だけでは足りません。** ご指摘のとおり、分析して示唆を出すのは新しく考えることを含みます。そこで役割を3つに切り分けることを提案します（4章）：A 整理／B 分析／C 新規発案。',
    '**最初の実行は IS**。IS分のAI分析用集計表（個人情報なし・全19シート）を作りました（6章）。スペースクリエイト（SC）の候補も出し、置き場は「採否待ち」にします。',
    '**GEOの各設問からのAI分析は可能です。** ISの22問それぞれについて、質問・AI回答の全文・回答に出た企業名を集計表に入れました。定型文は現状のまま参考として残し、候補の内容はAIの分析にします。',
    '**ROIとDSPは、AIが結果から施策を提案します。** DSPの判定項目（セッション・LPの直帰率・LPの滞在時間）のうち、**直帰率だけはデータにありません**（取得項目の追加が必要）。当面は滞在時間（エンゲージメント時間）と次ページ遷移率で代替します。',
    '**根拠はすべて「弱い」で出します**（確認済み）。ROIと月次レポートの数字は同じ出どころなので、根拠は1つと数えます（前回資料の「一致するので信頼できる」は撤回します）。',
    '**候補はPDCAダッシュボードの「採否待ち」に自動で表示されます**（最終形）。利用者が取り込む操作は要りません（3章 Q6）。',
]:
    a(B(t))

a(H(1, '2. 確認結果の反映'))
a(table([3200, 5000, 6370], ['項目', '中島さんの回答', '反映した内容'], [
    ['DSPの最新データ', '10月の実績が入った最新のダッシュボードを使う。DSP広告はSCHとISの2事業。配信データは配信会社作成のレポート（取込済み）', '最新ダッシュボード（2026-11-13更新・2026年10月分まで）を使用。IS分を集計表に反映。DSPはSCH・ISのみを対象にする'],
    ['HubSpotのレポート', '各事業の実績ファイルがCVリスト。実績ファイルを使う', 'IS｜実績.xlsm（2026-11-24更新）を件数だけ集計（氏名・メール・会社名は読まない）'],
    ['競合の社名', 'AIに渡してOK', 'GEOの回答に出た企業名を、そのまま集計表に入れた'],
    ['根拠の強さ', 'すべて「弱い」でOK', 'すべての候補を「弱い」で出す。根拠には、シート名・期間・数字を必ず書く'],
    ['最初の実行範囲', 'まず1事業。ISから', 'IS分の集計表を作成。SCも候補を出し、PDCAの施策候補を出すタブ（採否待ち）に置く'],
    ['定型文', 'いったんそのままの表現。内容はAIの分析からの示唆にする', '定型文は参考として残し、候補の内容は設問ごとのAI分析にする（Q3）'],
    ['AIの役割', '分析と示唆は新しく考えることではないのか。ROI・DSP・GEO・月次レポートでは不足。新たに考えたものも必要', '役割をA・B・Cに切り分ける提案（4章）'],
    ['DSPの判定項目', '目的は認知アップ・エンゲージメントアップ。セッション数、LPの直帰率の低下、LPの滞在時間の延長が閾値', '現状データと不足を整理（Q5）'],
    ['資料の形式', 'mdでは表が見にくい。Wordにする', 'このWord版を作成（rev.1のmdは古い版）'],
]))

a(H(1, '3. 質問への回答'))
a(H(2, 'Q1　分析と示唆を出すのは、新しく考えることにならないのですか？'))
a(P('なります。rev.1の「新しく考えることより整理」という書き方は不正確でした。**「材料にある事実を整理する」ことと「材料から読み取って考える」ことと「材料にない案を新しく発案する」ことは別の仕事**で、扱い（根拠・出所の書き方・確認の厚み）を分けた方が安全です。4章で切り分けを提案します。'))
a(H(2, 'Q2　月次レポートの数値はROIダッシュボードの機能で作っています。施策候補を出すのに問題はありますか？'))
a(P('**候補の出し方に問題はありません。** ただし、**同じ数字を2つの根拠として数えない**ようにします。ROIと月次レポートの数字は同じ出どころなので、根拠の件数は1つです。月次レポートが独自に持つ情報は、①人が書いた所見・示唆、②コンテンツ・メルマガ個別の指標（開封率・クリック率・読了率）、③課題と施策の進捗、④配信会社の広告レポートです。AI分析では、数字ではなくこの4つを月次レポートの根拠として使います。'))
a(H(2, 'Q3　GEOの各設問の結果から、実施した方がよい施策をAIが提案できますか？'))
a(P('**できます。** ISの22問それぞれについて、質問文・AIの回答全文・回答に出た企業名（競合）・自社が出たかを集計表（03_GEO_設問別）に入れました。AIは、各設問で「どんな会社が、どんな理由で挙がっているか」「自社のサイトに足りない情報は何か」を読み取り、設問単位で施策案（どのページ・記事・事例を作るか）を出します。定型文（業者推薦型は「会社比較・選定ガイド」など）は参考として残しますが、候補の内容は、設問ごとの分析にします。'))
a(P('限界：**全問が1つのAIモデル（claude-sonnet-5）の結果**で、質問タイプごとの件数は1〜8件です。そのため、GEO由来の候補はすべて根拠「弱い」です。'))
a(H(2, 'Q4　ROI・DSPにはPDCA用の示唆を出す機能がありません。結果から施策を提案できますか？'))
a(P('**できます。** ROIは、目標との差・段階別の転換率（MQL→SAL→SQL）・月ごとの変動から論点を読み取ります。DSPは、配信（imp・クリック・CTR・CPC・PC/SP別・素材別）とサイト側（流入・LP・滞在・検索）から、認知アップ・エンゲージメントアップの観点で提案します。実績ファイル（CVリスト）は、CVの元（どのメルマガ・資料）と商談への進み方（SAL率・SQL率）を材料にします。'))
a(H(2, 'Q5　DSPの判定項目（セッション・LPの直帰率・LPの滞在時間）は使えますか？'))
a(table([2600, 3600, 4200, 4170], ['判定項目', 'いまのデータ', 'IS・直近の値', '状態'], [
    ['① セッション数', 'あり（GA4・チャネル別、DSP流入のLP別）', 'Displayのセッション：9月 2,451 → 10月 1,680', '使える'],
    ['② LPの直帰率の低下', '**なし**（ダッシュボード／GA4のデータに直帰率が入っていない）', '—', '**取得項目の追加が必要**（GA4側）。当面は「次ページ遷移率」で代替'],
    ['③ LPの滞在時間の延長', 'エンゲージメント時間はあり（滞在時間そのものではない）', 'DSP流入の1セッションあたり：9月 9.7秒 → 10月 11.2秒', '使える（エンゲージメント時間で）'],
    ['（参考）次ページ遷移率', '月次レポート（9月）に記載', 'DSP広告からTOPページ流入後：約15%', '数値は記載のみ。月次の取得元は未確認'],
]))
a(P('閾値の具体的な数値は未決です。案：固定の数値は決めず、まず「前月比」と「直近3か月平均との差」で「改善／横ばい／悪化」を判定し、運用しながら数値を決めます。数値が必要な場合は、集計表（09_DSP判定の材料）の推移を見て決めてください。'))
a(H(2, 'Q6　「事業ごとの候補ファイル」「候補を取り込む」とは何ですか？ ダッシュボードに自動で表示されますよね？'))
a(P('**ご認識のとおりです。** 候補は **PDCAダッシュボード内（シートの「⏳ 採否待ち」）に表示**されます。最終形では、**毎営業日のバッチが自動で候補を追加**するので、利用者が「取り込む」操作をする必要はありません。「取り込む」という言い方は、裏方の説明が先に出てしまい、分かりにくくしました。'))
a(P('裏方の流れ（共通資料の決まりに沿ったもの）：'))
for t in [
    'AIの分析は、**中島さんが Claude Code で実行**します（バッチにはAIを使わない決まり）。',
    'その結果を「候補ファイル」（AIが作った候補の一覧）として、Boxの所定のフォルダに置きます。これはダッシュボードに候補を渡すための中間ファイルで、利用者は見ません。',
    '毎営業日のバッチが、前日までの候補ファイルを「まとめたデータ」に候補行として追加します。',
    '画面を開くと、「採否待ち」に自動で出ます。',
]:
    a(B(t))
a(P('バッチがまだ無いため、それまでの暫定をどうするかが次の検討です。案：アプリを開いたときに、設定済みの保存先フォルダの候補ファイルを**自動で読み込む**（画面の操作は不要）。「候補を取り込む」ボタン付きの画面は作らない方針に変更します（ご確認ください）。'))

a(H(1, '4. AIの役割の切り分け（提案）'))
a(P('候補の「出所」と「根拠の書き方」を、次の3区分で分けます。**根拠の強さはすべて「弱い」**（確認済み）。区分によって、根拠に書く内容と、人が確認する厚みが違います。'))
a(table([1700, 3300, 2400, 3770, 3400], ['区分', '内容', '出所の表記（案）', '根拠に書くこと', '歯止め'], [
    ['A 整理', '材料にある事実・人の所見を、施策の候補の形に整える（新しい内容は足さない）', '月次レポート／ROI／GEO／DSP', '元の数字・期間・件数・出どころ（シート名・スライド番号）', '元の材料と照合できる。AIが内容を足したら区分Bにする'],
    ['B 分析', '複数の材料を突き合わせて読み取った示唆（原因の仮説・注目すべき差）を、施策の候補にする', 'AI分析', '使った数字（どの表の、いつの、いくつ）と、読み取りの筋道（なぜその施策か）', '仮説と事実を分けて書く。使った数字が表に無ければ候補にしない'],
    ['C 新規発案', '材料に直接の裏付けがない、新しい施策案（他社・一般的な知見・想像に基づく案）を含む', 'AI発案', '「仮説」と明記。参考にした数字・事例。期待する指標と検証方法（検証指標・検証日の案）', '既存施策との重複を確認して記載。1事業あたりの件数に上限（案：5件）。採否で必ず人が検討'],
]))
a(P('共通の約束：'))
for t in [
    '渡された数字以外を根拠にしない（Cの「仮説」は、そのように明記する）。',
    '候補には、種類（新規／次サイクル）・推奨担当・推奨検証日・検証指標の案を付ける。',
    '個人情報・顧客名は渡さない（競合の社名は渡してよい）。',
    'AIが施策を確定しない。採否は人が理由つきで決める（アプリで、根拠が空の候補は採用できない）。',
    'アプリの「出所」の選択肢に、「月次レポート」「AI発案」を追加する（次の実装）。',
]:
    a(B(t))

a(H(1, '5. 材料の一覧（更新）'))
a(table([1900, 4000, 4300, 4370], ['材料', '場所・日付', '中身', '注意'], [
    ['ROI', '社内ドキュメント marketing_roi_2026_2026-11-09.json', '事業別・月別の目標と実績（セッション・MQL・SAL・SQL・受注・コスト・コンテンツ本数）', '2026-10まで入力済み。11月以降の「0」は未入力。MQL＝「cv」。月次レポートの数字と同じ出どころ'],
    ['月次レポート', '【外部】FLUED様フォルダ「FLUED｜2026｜月次定例レポート.pptx」（11/22更新）', '事業別の達成率・推移・コンテンツ評価・課題と施策。**人が書いた所見・示唆を含む**', '顧客名・担当者名を含む。AIには社名を除いて渡す。【外部】フォルダに書き戻さない'],
    ['GEO', 'GEO insights／results（2026-11-04実行ほか）', '設問ごとの質問・AI回答・出現の有無・回答に出た企業名', '1モデルのみ。質問の件数が少ない'],
    ['DSP広告・サイト', 'DSP広告分析ダッシュボード（2026-11-13更新・10月分まで）', '配信実績（月次・PC/SP別・素材別）、GA4（チャネル・LP別・ページ別）、GSC（検索）', 'DSP広告はSCH・ISの2事業。**LPの直帰率はデータなし**'],
    ['実績ファイル', '社内ドキュメント「マーケティング本部｜IS｜実績.xlsm」（11/24更新）ほか各事業', 'CVリスト（CVの元・流入チャネル・CV年月・SAL／SQL／受注）', '氏名・メール・会社名を含む。**件数だけ集計し、個人情報は出さない**'],
    ['既存施策', '施策管理表（正本）', '施策の件数（重複を避けるため）', '施策内容には事例企業名を含むため、件数のみ'],
]))

a(H(1, '6. IS分の集計表と、わかったこと'))
a(H(2, '6-1. 集計表の中身（個人情報なし・全19シート）'))
a(table([3300, 11270], ['シート', '中身'], [
    ['00 概要', '目的・シート一覧・根拠の扱い・AIの役割の切り分け'],
    ['01 ROI月次', 'セッション・MQL・SAL・SQL・受注・コンテンツ本数と目標（2024-02〜2026-10）、2026年度累計'],
    ['02 月次レポート所見', '2026年10月号のISの所見・示唆（社名除外）13項目。出典のスライド番号つき'],
    ['03 GEO設問別／03b 推移／04 定型文', 'ISの22問の質問・回答全文・出現・出た企業名／過去16回の実行の出現推移／定型文（現状のまま）'],
    ['05 DSP広告（月次／PC・SP別／素材別）', 'imp・クリック・CTR・CPC・費用・月予算（2025-11〜2026-10）'],
    ['06 サイト（GA4チャネル／DSP流入のLP別／ページ別／GSC月次／検索クエリ）', 'セッション・エンゲージメント時間・検索の表示回数とクリック、指名/非指名の検索クエリ'],
    ['07 実績CV（月×タイプ／月×流入チャネル／商談への進み方／CVの元）', 'CV件数とSAL・SQL・受注への進み方。メルマガ番号・CTA別の件数'],
    ['08 既存施策の件数 ／ 09 DSP判定の材料', '施策種類×ステータスの件数／DSPのセッション・滞在時間'],
]))
a(H(2, '6-2. 分かったこと（事実）と、見込まれる候補の例'))
a(P('候補の例は、材料から素直に読み取れる論点の**例**です（結論ではありません）。区分はA整理／B分析／C新規発案。', color='555555'))
a(table([2300, 6900, 5370], ['観点', '事実（出どころ）', '見込まれる候補の例（区分）'], [
    ['ROI（累計3〜10月）', 'セッション 17,578（目標比105%）、MQL 114（62%）、SQL 9（112%）。**SQL÷SAL は32%**（UN 87%・SCH 86%・MED 69%）。（ROI）', 'SALからSQLへの転換を上げる施策（B）。月次レポートの既存施策No.178〜180の次サイクル（A）'],
    ['9月の突出（MQL43・SAL19）', '9月のCV43件の内訳：メルマガ経由32・研究会10・フォーム入力1。**研究会経由は10件のうち9件がSAL（90%）、3件がSQL**。（実績ファイル）', '研究会（デジタルワークシート）のような接点の拡大（B）。9月のメルマガ32件の再現条件の確認（B）'],
    ['CVの元と商談への進み方', 'メルマガ経由231件（SAL率19.9%・SQL1率4.8%）、サービス資料DL 95件（SAL率11.6%）、お問い合わせ完了26件（SAL率19.2%・SQL1率15.4%）。（実績ファイル）', 'メルマガCTA別のSAL率の差を使った、配置・文言の見直し（B）'],
    ['DSP広告', '10月：imp 1,757,429・クリック3,999・CTR 0.23%・CPC約50円・費用201,205円。**CPCは3〜6月の約230〜310円から、7月以降は50〜96円に低下**（素材名の日付が20260701）。10月のPC CTR 0.05% に対しSP 1.03%。（DSPダッシュボード）', 'PC素材の差し替え（月次レポートも言及）（A）。SP中心の配信設計の検討（B）'],
    ['DSP流入後の動き', 'DSP流入のLPはほぼTOPページ（10月 1,611/1,642セッション）。1セッションあたりのエンゲージメント 9.7秒（9月）→ 11.2秒（10月）。DSP→TOP後の次ページ遷移率 約15%（月次レポート）。', '課題訴求型のコピー・強いヘッドラインのLPの検討（A：月次レポートの検討事項）。LPの回遊改善（B）'],
    ['GEO', 'IS 22問。業者推薦0/8・事例0/6・課題解決0/4・比較0/1・トレンド0/1。出たのは「会社名を直接聞く質問」だけで、8/31以降の7回は1〜2/2。（GEO）', '設問ごとの分析から、比較・選定ガイド／事例／課題起点の記事の案（B・C）。**根拠は弱い（1モデル・件数少）**'],
    ['月次レポートの示唆', '10月号に人が書いた示唆が3点（担当者のペインポイントを突くコンテンツの拡充／メルマガとコラムのマッチ度の見直し／コラム読了率の向上）。「検討の上、施策管理表への起票」と記載。', '3点をそのまま候補に（A）'],
    ['認知・流入', '月次レポート：Organic流入が目標に対し最も進捗率が低く、指名検索のクリックは0（9月）。', 'リスティングの再導入、非指名KWのコンテンツ（月次レポートの検討事項）（A）／C区分の新規案'],
]))

a(H(1, '7. 定型文（GEOの改善案）：現状のまま'))
a(P('「いったん現状のまま」とのご回答のため、そのまま掲載します。候補の内容は、3章Q3のとおり設問ごとのAI分析にします。'))
a(table([2000, 8400, 4170], ['質問タイプ', '定型文（現状のまま）', 'ISでの該当'], [
    ['業者推薦型', '候補として想起されるための「会社比較・選定ガイド」ページを整備。自社の強みを第三者視点の一覧に載る形（実績数・対応業種・ロット）で明示する。', '0/8'],
    ['事例・実績型', '導入事例ページを増やす。企業名・課題・成果を構造化データつきで掲載する。', '0/6'],
    ['課題解決型', '課題起点のノウハウ記事（○○の解決方法）を作り、その中で自社サービスを解決策として提示する。', '0/4'],
    ['ABM横断型', '対象意図に直接答える専用ページを用意する。', '（ISの設問にはなし）'],
]))

a(H(1, '8. 妥当性の担保（実施済み）'))
for t in [
    '候補の「根拠」に、元の数字・期間・件数・出どころを書く。**根拠が空の候補は採用できない**（アプリに実装済み）。',
    '「根拠の強さ」（十分／弱い：件数が少ない／弱い：データが古い／弱い：その他）を明記する。**今回はすべて「弱い」**。弱い候補は黄色表示で、採用時に確認が出る。',
    '候補には、出所（区分）・提案日・根拠を記入する欄がある（アプリに実装済み）。',
]:
    a(B(t))

a(H(1, '9. 残りの確認事項と次の手順'))
a(H(2, '確認していただきたいこと'))
for t in [
    '4章の切り分け（A整理／B分析／C新規発案）と、Cの件数上限（案：1事業5件）でよいか。',
    'Q6の方針：候補の追加は自動（バッチ、暫定はアプリ起動時の自動読込）とし、「候補を取り込む」画面は作らない。',
    'DSPのLPの直帰率：GA4の取得項目に追加してよいか（Apps Script／ダッシュボードの変更を伴う）。',
    'SCの候補の置き場：「採否待ち」にSCを5つ目の事業として加える案でよいか。',
]:
    a(B(t))
a(H(2, '次の手順（別チャットで実施）'))
for t in [
    '（完了）AIに渡す集計表をIS分で作成。',
    'AI分析（IS）を実行し、候補ファイルを作る（区分A・B・C、すべて根拠「弱い」）。',
    '候補を「採否待ち」に自動で出す仕組み（バッチ、または暫定の自動読込）。SC用の置き場を含む。',
    '中島さんが候補を確認し、採否の流れを試す。問題なければ、SCH・MED・UN・SC分の集計表を作る。',
]:
    a(B(t))

footer_xml = (f'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
              f'<w:r>{rpr(16, False, "777777")}<w:t xml:space="preserve">AI分析の判断材料 確認資料（rev.2）　</w:t></w:r>'
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
core = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>AI分析の判断材料 確認資料（rev.2）</dc:title><dc:creator>KPI-PDCA</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-11-25T00:00:00Z</dcterms:created></cp:coreProperties>'''

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
