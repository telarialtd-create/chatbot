#!/usr/bin/env python3
"""
スケジュールシートの今週分をLINEグループに送信
GitHub Actions から毎日 JST 03:00 に実行される（Mac不要）
"""
import json
import os
import subprocess
import tempfile
import time
import urllib.request
import urllib.parse
from datetime import date, timedelta, datetime, timezone

JST = timezone(timedelta(hours=9))

SHEET_ID   = "10siqLe6B9A7uvNWgRUdHb462RqxCxkGEGMEKTPhY-S8"
LINE_TOKEN  = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_GROUP_ID = os.environ.get("LINE_GROUP_ID", "C98b65635d31c24893cf8c0fc61070065")

ROW_START = 1
WEEK_COLS = 7


def col_num_to_letter(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def http_get_with_retry(url, headers, max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            req = urllib.request.Request(url)
            for k, v in headers.items():
                req.add_header(k, v)
            with urllib.request.urlopen(req, timeout=60) as resp:
                return resp.read()
        except Exception as e:
            print(f"    [HTTP] 試行{attempt}/{max_attempts} 失敗: {e}")
            if attempt < max_attempts:
                time.sleep(30)
    raise RuntimeError(f"HTTP取得に{max_attempts}回失敗: {url[:80]}")


def get_access_token():
    """環境変数 → ローカルファイルの順で認証情報を取得"""
    client_id     = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    refresh_token = os.environ.get("GOOGLE_REFRESH_TOKEN")

    if not client_id:
        with open(os.path.expanduser("~/.config/gcp-oauth.keys.json")) as f:
            keys = json.load(f)["installed"]
        with open(os.path.expanduser("~/.config/gdrive-server-credentials.json")) as f:
            creds = json.load(f)
        client_id     = keys["client_id"]
        client_secret = keys["client_secret"]
        refresh_token = creds["refresh_token"]

    data = urllib.parse.urlencode({
        "client_id":     client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type":    "refresh_token",
    }).encode()

    for attempt in range(1, 4):
        try:
            req = urllib.request.Request(
                "https://oauth2.googleapis.com/token", data=data, method="POST")
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read())["access_token"]
        except Exception as e:
            print(f"    [トークン取得] 試行{attempt}/3 失敗: {e}")
            if attempt < 3:
                time.sleep(30)
    raise RuntimeError("アクセストークン取得に3回失敗")


def get_sheet_info(access_token):
    today  = datetime.now(JST).date()
    target = f"{today.year}年{today.month}月"
    url    = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}?fields=sheets.properties"
    raw    = http_get_with_retry(url, {"Authorization": f"Bearer {access_token}"})
    data   = json.loads(raw)
    sheets = {s["properties"]["title"].strip(): (str(s["properties"]["sheetId"]), s["properties"]["title"].strip())
              for s in data["sheets"]}
    recent = [t for t in sheets if str(today.year) in t]
    print(f"    利用可能シート(今年): {recent[:10]}")
    if target in sheets:
        gid, title = sheets[target]
        print(f"    シート検出: {title}")
        return gid, title
    for m in range(today.month - 1, 0, -1):
        key = f"{today.year}年{m}月"
        if key in sheets:
            gid, title = sheets[key]
            print(f"    シート検出(フォールバック): {title} ※{target}が見つからなかった")
            return gid, title
    raise RuntimeError(f"シートが見つかりません: {target}")


def get_date_column(access_token, sheet_title):
    today   = datetime.now(JST).date()
    encoded = urllib.parse.quote(sheet_title)
    url     = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/"
               f"{encoded}!1:1?majorDimension=ROWS")
    raw     = http_get_with_retry(url, {"Authorization": f"Bearer {access_token}"})
    values  = json.loads(raw).get("values", [[]])[0]
    matches = [i + 1 for i, v in enumerate(values) if str(v).strip() == str(today.day)]
    if not matches:
        raise RuntimeError(f"行1に {today.day} 日が見つかりません")
    sheet_month = int(sheet_title.replace("年", "月").split("月")[1]) if "年" in sheet_title else today.month
    return matches[0] if sheet_month == today.month else matches[-1]


def get_last_row(access_token, gid):
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/A1:A200"
           f"?majorDimension=ROWS")
    try:
        raw  = http_get_with_retry(url + f"&gid={gid}", {"Authorization": f"Bearer {access_token}"})
        data = json.loads(raw)
    except Exception:
        data = {"values": []}
    for i, row in enumerate(data.get("values", []), 1):
        if row and "講習" in row[0]:
            return i
    return 108


def get_range(access_token, gid, sheet_title):
    start_col = get_date_column(access_token, sheet_title)
    end_col   = start_col + WEEK_COLS - 1
    last_row  = get_last_row(access_token, gid)
    start     = f"{col_num_to_letter(start_col)}{ROW_START}"
    end       = f"{col_num_to_letter(end_col)}{last_row}"
    print(f"    今日の列: {col_num_to_letter(start_col)}（列{start_col}）→ {start}:{end}")
    return f"{start}:{end}", last_row, start_col, end_col


def export_range_as_png(access_token, range_str, out_path, gid):
    params = urllib.parse.urlencode({
        "format": "pdf", "range": range_str, "gid": gid,
        "portrait": "true", "size": "6", "fith": "true",
        "top_margin": "0", "bottom_margin": "0", "left_margin": "0", "right_margin": "0",
        "sheetnames": "false", "printtitle": "false", "pagenumbers": "false",
        "gridlines": "true", "notes": "false",
    })
    url      = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?{params}"
    pdf_data = http_get_with_retry(url, {"Authorization": f"Bearer {access_token}"})

    pdf_path = out_path.replace(".png", ".pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_data)

    import fitz
    from PIL import Image
    import numpy as np

    doc       = fitz.open(pdf_path)
    page_imgs = []
    for page in doc:
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        gray = np.array(img.convert("L"))
        dark_per_row = (gray < 80).sum(axis=1)
        if (dark_per_row > img.width * 0.15).any():
            page_imgs.append((img, dark_per_row))
    doc.close()

    if not page_imgs:
        raise RuntimeError("有効なページが見つかりません")

    if len(page_imgs) == 1:
        combined = page_imgs[0][0]
    else:
        strips = [page_imgs[0][0]]
        for img, dark_per_row in page_imgs[1:]:
            top_rows = dark_per_row[:min(200, len(dark_per_row))]
            thresh   = max(top_rows) * 0.8 if max(top_rows) > 0 else img.width * 0.5
            header_end = 0
            for y in range(len(top_rows)):
                if top_rows[y] >= thresh:
                    header_end = y + 1
                    break
            strips.append(img.crop((0, header_end, img.width, img.height)))
        total_h  = sum(i.height for i in strips)
        max_w    = max(i.width for i in strips)
        combined = Image.new("RGB", (max_w, total_h), "white")
        y = 0
        for img in strips:
            combined.paste(img, (0, y))
            y += img.height

    combined.save(out_path)
    return out_path


def download_sheet_as_png(access_token, range_str, gid, last_row, sheet_title, start_col, end_col):
    from PIL import Image

    tmp_dir     = tempfile.mkdtemp()
    week_col    = col_num_to_letter(start_col)
    week_end    = col_num_to_letter(end_col)
    left_range  = f"A1:C{last_row}"
    right_range = f"{week_col}1:{week_end}{last_row}"

    left_png  = export_range_as_png(access_token, left_range,  os.path.join(tmp_dir, "left.png"),  gid)
    right_png = export_range_as_png(access_token, right_range, os.path.join(tmp_dir, "right.png"), gid)

    def crop_ws(img, pad=4):
        gray = img.convert("L")
        mask = gray.point(lambda p: 0 if p >= 253 else 255)
        bbox = mask.getbbox()
        if bbox:
            w, h = img.size
            return img.crop((max(0, bbox[0]-pad), max(0, bbox[1]-pad),
                             min(w, bbox[2]+pad), min(h, bbox[3]+pad)))
        return img

    left_img  = crop_ws(Image.open(left_png))
    right_img = crop_ws(Image.open(right_png))
    H = min(left_img.height, right_img.height)
    combined  = Image.new("RGB", (left_img.width + right_img.width, H), "white")
    combined.paste(left_img.crop((0, 0, left_img.width, H)),   (0, 0))
    combined.paste(right_img.crop((0, 0, right_img.width, H)), (left_img.width, 0))

    out_path = os.path.join(tmp_dir, "combined.jpg")
    combined.save(out_path, "JPEG", quality=90)
    return out_path


def upload_image(image_path):
    services = [
        lambda: subprocess.run(
            ["curl", "-s", "-F", "reqtype=fileupload", "-F", f"fileToUpload=@{image_path}",
             "https://catbox.moe/user/api.php"],
            capture_output=True, text=True).stdout.strip(),
        lambda: subprocess.run(
            ["curl", "-s", "-F", "reqtype=fileupload", "-F", "time=72h",
             "-F", f"fileToUpload=@{image_path}",
             "https://litterbox.catbox.moe/resources/internals/api.php"],
            capture_output=True, text=True).stdout.strip(),
    ]
    for fn in services:
        try:
            url = fn()
            if url.startswith("https://"):
                return url
        except Exception:
            continue
    raise RuntimeError("全アップロードサービスが失敗")


def send_to_line(image_url):
    data = json.dumps({
        "to": LINE_GROUP_ID,
        "notificationDisabled": True,
        "messages": [{"type": "image",
                      "originalContentUrl": image_url,
                      "previewImageUrl":    image_url}],
    }).encode()
    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/push", data=data,
        headers={"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def main():
    import datetime
    print(f"[開始] {datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')} UTC")

    print("[1] アクセストークン取得中...")
    access_token = get_access_token()

    gid, sheet_title = get_sheet_info(access_token)
    range_str, last_row, start_col, end_col = get_range(access_token, gid, sheet_title)

    print("[2] スプレッドシートをPDF→PNG変換中...")
    image_path = download_sheet_as_png(access_token, range_str, gid, last_row,
                                       sheet_title, start_col, end_col)
    print(f"    保存先: {image_path}")

    print("[3] 画像アップロード中...")
    image_url = upload_image(image_path)
    print(f"    URL: {image_url}")

    print("[4] LINE送信中...")
    result = send_to_line(image_url)
    print(f"[完了] {result}")


if __name__ == "__main__":
    main()
