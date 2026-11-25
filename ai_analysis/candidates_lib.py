# -*- coding: utf-8 -*-
"""候補ファイルの書き出しの共通部品（make_candidates*.py から使う）。

- 候補ファイル：data/candidates/candidates_<事業>_<日付>.json（gitに置く。個人情報は含まない）
- 一括読込用の索引：data/candidates/candidates.js（`window.KPI_CANDIDATES=[…];`）。全候補ファイルをまとめたもの。
  アプリは、HTMLと同じ場所の data/candidates/candidates.js を <script> で読むので、
  ダブルクリックで開いても（フォルダの許可なしで）候補が出る。候補ファイルを足したら、必ず索引を作り直す。
"""
import os, json, glob, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAND_DIR = os.path.join(ROOT, 'data', 'candidates')
TODAY = '2026-11-25'


def check_items(items):
    ids = [c['originId'] for c in items]
    assert len(ids) == len(set(ids)), 'originId の重複'
    assert sum(1 for c in items if c['cls'] == 'C新規発案') <= 5, 'C区分は5件まで'
    for c in items:
        assert c['evidence'].strip(), c['originId'] + ' 根拠が空'
        assert c['evidenceLevel'].startswith('弱い'), c['originId']
        assert c['naiyo'].strip() and c['mokuteki'].strip(), c['originId']
        assert c['cls'][0] in 'ABC', c['originId']


def write_candidates(biz, items, src, date=None, note=None):
    """候補ファイルを書き、索引を作り直す。書いたファイルのパスを返す。"""
    check_items(items)
    date = date or TODAY
    os.makedirs(CAND_DIR, exist_ok=True)
    cands = []
    for c in items:
        d = dict(c)
        d['biz'] = biz
        cands.append(d)
    data = {
        'schemaVersion': '0.1',
        'kind': 'candidates',
        'generatedAt': datetime.datetime.now().astimezone().isoformat(timespec='seconds'),
        'generatedBy': 'AI分析（Claude Code。中島さんが実行）',
        'biz': biz,
        'source': src,
        'note': note or '根拠の強さはすべて「弱い」。区分A＝整理／B＝分析／C＝新規発案（仮説。1事業5件まで）。採否は人が理由つきで決める。',
        'candidates': cands,
    }
    path = os.path.join(CAND_DIR, 'candidates_%s_%s.json' % (biz, date.replace('-', '')))
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    rebuild_index()
    return path


def rebuild_index():
    files = sorted(glob.glob(os.path.join(CAND_DIR, 'candidates_*.json')))
    docs = [json.load(open(f, encoding='utf-8')) for f in files]
    body = ('/* 自動生成（ai_analysis/candidates_lib.py）。直接編集しない。候補ファイル（candidates_*.json）をまとめた索引。\n'
            '   アプリが <script> で読み、状態「候補」の行として該当事業のシートに追加する。 */\n'
            'window.KPI_CANDIDATES=' + json.dumps(docs, ensure_ascii=False, separators=(',', ':')) + ';\n')
    with open(os.path.join(CAND_DIR, 'candidates.js'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(body)
    return files
