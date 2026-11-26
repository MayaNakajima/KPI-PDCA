# -*- coding: utf-8 -*-
"""Teams 通知のテスト送信（KPI-PDCA）

Teams の「ワークフロー」で作った Webhook（要求を受け取ったらチャット／チャネルに投稿する）へ、
テスト用のカード（Adaptive Card）を1件送ります。

- Webhook の URL は「URLを知っていれば誰でも送れる」ため、APIキーと同じ扱いにします。
  リポジトリ・HTML・ログ・チャットには書かず、閲覧を限定した Box のファイルに1行で保存し、
  そのファイルの場所を --url-file で渡します。
- この処理は URL を画面に表示しません。
- 標準ライブラリだけで動きます。

使い方（PowerShell）:
  & C:\\work\\anaconda_install\\python.exe -I batch\\notify_test.py --url-file "<URLを保存したファイル>"
  & C:\\work\\anaconda_install\\python.exe -I batch\\notify_test.py --dry-run   # 送らずに中身だけ表示

通知の文言を変えるとき（Teams のワークフロー側は、受け取ったカードをそのまま投稿する設定にしておく。文言はこちらで決める）:
  --title "見出し"  --text "本文1行目" --text "本文2行目"   … 文言を指定する（--text は何回でも）
  --message-file 文言.txt                                   … 1行目＝見出し、2行目以降＝本文（UTF-8）
"""
import argparse
import datetime
import json
import sys
import urllib.error
import urllib.request


def build_card(title, lines):
    body = [
        {"type": "TextBlock", "size": "Large", "weight": "Bolder", "wrap": True, "text": title},
        {"type": "TextBlock", "isSubtle": True, "wrap": True,
         "text": datetime.datetime.now().strftime("%Y/%m/%d %H:%M") + "　送信元：KPI-PDCA バッチ（テスト）"},
    ]
    for line in lines:
        body.append({"type": "TextBlock", "wrap": True, "text": line})
    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": body,
    }


def build_payload(card):
    # Teams のワークフロー（Webhook 受信）が受け取る形：カードを attachments に入れたメッセージ
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": None,
            "content": card,
        }],
    }


def read_url(path):
    with open(path, encoding="utf-8-sig") as f:
        url = f.read().strip()
    if not url.lower().startswith("https://"):
        raise ValueError("ファイルの1行目に https:// で始まる URL を入れてください")
    return url


def main():
    ap = argparse.ArgumentParser(description="Teams 通知のテスト送信")
    ap.add_argument("--url-file", help="Webhook URL を1行で保存したファイルの場所")
    ap.add_argument("--dry-run", action="store_true", help="送らずに送信内容を表示する")
    ap.add_argument("--title", help="通知の見出し（省略すると、テスト送信の見出し）")
    ap.add_argument("--text", action="append", help="本文の1行（何回でも指定できる）")
    ap.add_argument("--message-file", help="文言のファイル（UTF-8。1行目＝見出し、2行目以降＝本文）")
    args = ap.parse_args()

    title = "【KPI-PDCA通知】テスト送信"
    lines = [
        "これは通知のテストです。このメッセージが届いていれば、設定は完了しています。",
        "本番では、停滞（採否待ち・検証日が空・実績が空・判定が空・次サイクルなし）の件数と対象の施策をお知らせします。",
    ]
    if args.message_file:
        with open(args.message_file, encoding="utf-8-sig") as f:
            rows = [l.rstrip("\r") for l in f.read().split("\n") if l.strip()]
        if rows:
            title, lines = rows[0], rows[1:]
    if args.title:
        title = args.title
    if args.text:
        lines = args.text
    card = build_card(title, lines)
    payload = build_payload(card)

    if args.dry_run:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    if not args.url_file:
        print("--url-file を指定してください（または --dry-run）", file=sys.stderr)
        return 2

    try:
        url = read_url(args.url_file)
    except (OSError, ValueError) as e:
        print("URL ファイルを読めませんでした：" + str(e), file=sys.stderr)
        return 2

    req = urllib.request.Request(
        url, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            print("送信しました（HTTP %d）。Teams に届いたか確認してください。" % res.status)
            return 0
    except urllib.error.HTTPError as e:
        print("送信に失敗しました（HTTP %d）。ワークフローの設定を確認してください。" % e.code, file=sys.stderr)
        return 1
    except urllib.error.URLError as e:
        print("送信先に接続できませんでした：" + str(e.reason), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
