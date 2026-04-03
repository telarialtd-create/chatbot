#!/usr/bin/env python3
import urllib.request
import re
import json
import os
from datetime import datetime

LINE_TOKEN = os.environ["LINE_TOKEN"]
LINE_USER_ID = os.environ["LINE_USER_ID"]

TARGETS = ["CREA", "ふわもこ"]

def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", errors="ignore")

def find_rank(pairs, target):
    for name, rank in pairs:
        if target in name:
            return rank
    return "圏外"

def parse_estama_access():
    html = fetch("https://estama.jp/ranking/6/store/?area_m_id=34&genre_id=2")
    shops = re.findall(r'main_details_shop_name.*?<a[^>]*>([^<]+)</a>', html, re.DOTALL)
    return [(s.strip(), i) for i, s in enumerate(shops, 1)]

def parse_estama_omotenashi():
    html = fetch("https://estama.jp/etc/ranking/service/3400/")
    shops = re.findall(r'main_details_shop_name.*?<a[^>]*>([^<]+)</a>', html, re.DOTALL)
    return [(s.strip(), i) for i, s in enumerate(shops, 1)]

def parse_eslove():
    html = fetch("https://eslove.jp/chugoku-shikoku/hiroshima/shoplist/ranking/")
    matches = re.findall(r"data-gtm-rank\':\'(\d+)\'.*?data-gtm-shopname\':\'([^\']+)\'", html)
    return [(name, int(rank)) for rank, name in matches]

def parse_esthe_ranking():
    html = fetch("https://www.esthe-ranking.jp/hiroshima/hiroshima-city/ippan/")
    pairs = []
    top3 = re.findall(r'alt="(\d+)位"[^>]*>.*?<b>([^<]+)</b>', html, re.DOTALL)
    for rank, name in top3:
        pairs.append((name.strip(), int(rank)))
    rest = re.findall(r'<span class="dropcap-bg">(\d+)位</span>.*?<b>([^<]+)</b>', html, re.DOTALL)
    for rank, name in rest:
        pairs.append((name.strip(), int(rank)))
    return pairs

def parse_ekichika():
    html = fetch("https://ranking-deli.jp/fuzoku/style8/32/")
    matches = re.findall(r'"position":(\d+),"url":"[^"]+","name":"([^"]+)"', html)
    return [(name, int(pos)) for pos, name in matches]

def get_rankings():
    sites = [
        ("エステ魂 アクセス", parse_estama_access),
        ("エステ魂 おもてなし", parse_estama_omotenashi),
        ("エステラブ", parse_eslove),
        ("メンエスランキング", parse_esthe_ranking),
        ("エキチカ", parse_ekichika),
    ]
    results = {}
    for site_name, parser in sites:
        try:
            pairs = parser()
            results[site_name] = {t: find_rank(pairs, t) for t in TARGETS}
        except Exception as e:
            results[site_name] = {t: "取得失敗" for t in TARGETS}
    return results

def build_message(results):
    today = datetime.now().strftime("%Y/%m/%d")
    lines = [f"📊 ランキング通知 {today}\n"]
    for shop in TARGETS:
        label = "CREA" if shop == "CREA" else "ふわもこSPA"
        lines.append(f"【{label}】")
        for site, ranks in results.items():
            r = ranks[shop]
            rank_str = f"{r}位" if isinstance(r, int) else r
            lines.append(f"  {site}：{rank_str}")
        lines.append("")
    return "\n".join(lines).strip()

def send_line(message):
    body = json.dumps({
        "to": LINE_USER_ID,
        "messages": [{"type": "text", "text": message}]
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/push",
        data=body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {LINE_TOKEN}"
        }
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode()

if __name__ == "__main__":
    results = get_rankings()
    message = build_message(results)
    print(message)
    send_line(message)
    print("送信完了")
