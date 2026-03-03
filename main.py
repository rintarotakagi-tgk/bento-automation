import os
import requests
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NOTION_API_KEY = os.environ["NOTION_API_KEY"]
NOTION_DATABASE_ID = os.environ["NOTION_DATABASE_ID"]
SLACK_WEBHOOK_URL = os.environ["SLACK_WEBHOOK_URL"]

TODAY = date.today().isoformat()
TODAY_LABEL = date.today().strftime("%Y年%m月%d日")


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


def create_excel(orders, filepath):
    wb = Workbook()
    ws = wb.active
    ws.title = "お弁当注文"

    # ページ設定（A4印刷）
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_margins.left = 0.75
    ws.page_margins.right = 0.75
    ws.page_margins.top = 1.0
    ws.page_margins.bottom = 1.0

    # カラー定義
    header_fill = PatternFill("solid", start_color="2F5496")
    alt_fill = PatternFill("solid", start_color="DCE6F1")
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # タイトル
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = f"お弁当注文リスト　{TODAY_LABEL}"
    title_cell.font = Font(name="Arial", bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.append([])  # 空白行

    # ヘッダー
    headers = ["No.", "注文者名", "注文内容", "✓"]
    ws.append(headers)
    header_row = 3
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col)
        cell.value = h
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[header_row].height = 22

    # データ行
    for i, order in enumerate(orders, 1):
        row_num = header_row + i
        fill = alt_fill if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        row_data = [i, order["name"], order["item"], ""]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col)
            cell.value = val
            cell.font = Font(name="Arial", size=11)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if col in [1, 4] else "left", vertical="center")
        ws.row_dimensions[row_num].height = 22

    # 集計行
    total_row = header_row + len(orders) + 1
    ws.merge_cells(f"A{total_row}:B{total_row}")
    total_cell = ws[f"A{total_row}"]
    total_cell.value = f"合計：{len(orders)} 名"
    total_cell.font = Font(name="Arial", bold=True, size=11)
    total_cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[total_row].height = 20

    # 列幅
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 8

    wb.save(filepath)
    print(f"Excel作成完了: {filepath}")


def post_to_slack(orders):
    lines = [f"🍱 *{TODAY_LABEL} お弁当注文リスト*（{len(orders)}名）\n"]
    for i, order in enumerate(orders, 1):
        lines.append(f"{i}. {order['name']}　→　{order['item']}")
    lines.append("\n📎 添付のExcelを印刷してチェックシートとしてご利用ください。")

    res = requests.post(SLACK_WEBHOOK_URL, json={"text": "\n".join(lines)})
    res.raise_for_status()
    print("Slack投稿完了")


def main():
    print(f"=== お弁当注文Slack通知 ({TODAY}) ===")

    orders = fetch_orders()
    print(f"注文件数: {len(orders)} 件")

    if not orders:
        requests.post(SLACK_WEBHOOK_URL, json={"text": f"🍱 {TODAY_LABEL} の注文はありませんでした。"})
        print("注文なし通知を送信しました。")
        return

    excel_path = f"/tmp/bento_{TODAY}.xlsx"
    create_excel(orders, excel_path)
    post_to_slack(orders)

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
