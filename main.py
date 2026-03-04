import os
import requests
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

NOTION_API_KEY = os.environ["NOTION_API_KEY"]
NOTION_DATABASE_ID = os.environ["NOTION_DATABASE_ID"]
SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]
SLACK_CHANNEL = os.environ["SLACK_CHANNEL"]

TODAY = date.today().isoformat()
TODAY_LABEL = date.today().strftime("%Y年%m月%d日")

# フォント登録（環境に合わせてパスを変更してください）
FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"
pdfmetrics.registerFont(TTFont("IPAGothic", FONT_PATH))


def fetch_orders():
    url = f"https://api.notion.com/v1/databases/{NOTION_DATABASE_ID}/query"
    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }
    payload = {
        "filter": {
            "property": "注文日",
            "date": {"equals": TODAY}
        },
        "sorts": [{"property": "注文者名", "direction": "ascending"}]
    }
    res = requests.post(url, headers=headers, json=payload)
    res.raise_for_status()
    results = res.json().get("results", [])

    orders = []
    for page in results:
        props = page["properties"]

        name_prop = props.get("注文者名", {})
        if name_prop.get("type") == "title":
            name = "".join([t["plain_text"] for t in name_prop.get("title", [])])
        else:
            name = "".join([t["plain_text"] for t in name_prop.get("rich_text", [])])

        item_prop = props.get("注文内容", {})
        item = "".join([t["plain_text"] for t in item_prop.get("rich_text", [])])

        if name:
            orders.append({"name": name, "item": item})

    return orders


def create_pdf(orders, filepath):
    w, h = A4
    c = canvas.Canvas(filepath, pagesize=A4)

    # 色定義
    header_bg = HexColor("#2F5496")
    alt_bg = HexColor("#DCE6F1")
    gray_bg = HexColor("#F2F2F2")
    border_color = HexColor("#AAAAAA")

    # 列定義（ページ中央配置）
    col_widths = [12 * mm, 45 * mm, 95 * mm, 12 * mm]
    table_width = sum(col_widths)
    margin_left = (w - table_width) / 2
    col_x = [margin_left]
    for cw in col_widths[:-1]:
        col_x.append(col_x[-1] + cw)

    row_height = 7 * mm
    header_height = 8 * mm
    title_height = 12 * mm

    y = h - 20 * mm

    # --- タイトル行 ---
    c.setFillColor(gray_bg)
    c.rect(margin_left, y - title_height, table_width, title_height, fill=1, stroke=0)
    c.setStrokeColor(header_bg)
    c.setLineWidth(1.5)
    c.rect(margin_left, y - title_height, table_width, title_height, fill=0, stroke=1)
    c.setFillColor(HexColor("#333333"))
    c.setFont("IPAGothic", 14)
    c.drawCentredString(margin_left + table_width / 2, y - title_height + 3.5 * mm, f"お弁当注文リスト　{TODAY_LABEL}")
    y -= title_height

    # --- ヘッダー行 ---
    c.setFillColor(header_bg)
    c.rect(margin_left, y - header_height, table_width, header_height, fill=1, stroke=0)
    c.setStrokeColor(border_color)
    c.setLineWidth(0.5)
    c.setFont("IPAGothic", 9)
    c.setFillColor(white)
    headers = ["No.", "注文者名", "注文内容", "✓"]
    for i, (hdr, cx, cw) in enumerate(zip(headers, col_x, col_widths)):
        c.drawCentredString(cx + cw / 2, y - header_height + 2.5 * mm, hdr)
        c.rect(cx, y - header_height, cw, header_height, fill=0, stroke=1)
    y -= header_height

    # --- データ行 ---
    c.setFont("IPAGothic", 9)
    for i, order in enumerate(orders, 1):
        fill = alt_bg if i % 2 == 0 else white
        c.setFillColor(fill)
        c.rect(margin_left, y - row_height, table_width, row_height, fill=1, stroke=0)
        c.setStrokeColor(border_color)
        c.setLineWidth(0.5)

        values = [str(i), order["name"], order["item"], ""]
        aligns = ["center", "left", "left", "center"]
        for val, align, cx, cw in zip(values, aligns, col_x, col_widths):
            c.rect(cx, y - row_height, cw, row_height, fill=0, stroke=1)
            c.setFillColor(HexColor("#333333"))
            text_y = y - row_height + 2 * mm
            if align == "center":
                c.drawCentredString(cx + cw / 2, text_y, val)
            else:
                c.drawString(cx + 2 * mm, text_y, val)
        y -= row_height

    # --- 合計行 ---
    total_height = 7 * mm
    c.setFillColor(gray_bg)
    c.rect(margin_left, y - total_height, table_width, total_height, fill=1, stroke=0)
    c.setStrokeColor(header_bg)
    c.setLineWidth(1.5)
    c.rect(margin_left, y - total_height, table_width, total_height, fill=0, stroke=1)
    c.setFillColor(HexColor("#333333"))
    c.setFont("IPAGothic", 9)
    total_text = f"合計：{len(orders)} 名"
    c.drawRightString(col_x[3] - 2 * mm, y - total_height + 2 * mm, total_text)

    c.save()
    print(f"PDF作成完了: {filepath}")


def upload_to_slack(filepath):
    filename = os.path.basename(filepath)
    headers = {"Authorization": f"Bearer {SLACK_BOT_TOKEN}"}

    res = requests.get(
        "https://slack.com/api/files.getUploadURLExternal",
        headers=headers,
        params={"filename": filename, "length": os.path.getsize(filepath)}
    )
    res.raise_for_status()
    data = res.json()
    if not data.get("ok"):
        raise RuntimeError(f"URL取得失敗: {data}")

    upload_url = data["upload_url"]
    file_id = data["file_id"]

    with open(filepath, "rb") as f:
        res = requests.post(upload_url, files={"file": f})
    res.raise_for_status()

    res = requests.post(
        "https://slack.com/api/files.completeUploadExternal",
        headers={**headers, "Content-Type": "application/json"},
        json={
            "files": [{"id": file_id}],
            "channel_id": SLACK_CHANNEL,
            "initial_comment": f"🍱 {TODAY_LABEL} のお弁当注文リストです。印刷してご使用ください。"
        }
    )
    res.raise_for_status()
    data = res.json()
    if not data.get("ok"):
        raise RuntimeError(f"投稿失敗: {data}")

    print("Slackへのファイル投稿完了")


def main():
    print(f"=== お弁当注文Slack通知 ({TODAY}) ===")

    orders = fetch_orders()
    print(f"注文件数: {len(orders)} 件")

    if not orders:
        requests.post(
            "https://slack.com/api/chat.postMessage",
            headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}", "Content-Type": "application/json"},
            json={"channel": SLACK_CHANNEL, "text": f"🍱 {TODAY_LABEL} の注文はありませんでした。"}
        )
        print("注文なし通知を送信しました。")
        return

    pdf_path = f"/tmp/bento_{TODAY}.pdf"
    create_pdf(orders, pdf_path)
    upload_to_slack(pdf_path)

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
