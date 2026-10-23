# -*- coding: utf-8 -*-
"""生成台账转录的 demo 样张（纯虚构数据）：
1. sample/sheets/register_sheet_001.png —— 仿手写“项目用印登记表”扫描样张
2. sample/sheets/ground_truth.json     —— 样张标注（供 --mock 离线演示）
3. sample/ledger_template.xlsx         —— 电子台账模板（待回写）
"""
import json
import os
import random

from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SHEET_DIR = os.path.join(HERE, "sample", "sheets")

TITLE = "项目用印登记表"
HEADERS = ["序号", "日期", "文件名称", "份数", "用章类型", "经办人"]
RECORDS = [
    ["1", "2026.09.02", "关于A栋幕墙样板段验收的函", "2", "公章", "王工"],
    ["2", "2026.09.05", "消防检测委托合同", "1", "合同章", "李工"],
    ["3", "2026.09.09", "景观铺装材料询价函", "3", "公章", "张工"],
    ["4", "2026.09.12", "电梯厅门到货确认单", "1", "公章", "王工"],
    ["5", "2026.09.16", "精装样板段施工方案报审表", "2", "公章", ""],   # 经办人漏填 → 演示校验拦截
    ["6", "2026.09.20", "幕墙深化设计意见回复函", "2", "公章", "李工"],
]

COL_W = [90, 210, 470, 130, 210, 210]
MARGIN_X, TOP_Y, ROW_H, HEADER_H = 90, 190, 115, 70
FONT_SIZES = {"title": 44, "header": 30, "value": 30}


def _font(size, handwriting=True):
    candidates = ["simkai.ttf", "STKAITI.TTF", "msyh.ttc", "simhei.ttf"] if handwriting \
        else ["simhei.ttf", "msyh.ttc"]
    for name in candidates:
        path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name)
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def make_sheet_image(path):
    rng = random.Random(42)
    width = MARGIN_X * 2 + sum(COL_W)
    height = TOP_Y + HEADER_H + ROW_H * (len(RECORDS) + 1) + 120
    img = Image.new("RGB", (width, height), (248, 246, 240))
    draw = ImageDraw.Draw(img)

    draw.text((width // 2, 90), TITLE, font=_font(FONT_SIZES["title"], False),
              fill=(20, 20, 20), anchor="mm")

    y0 = TOP_Y
    xs = [MARGIN_X]
    for w in COL_W:
        xs.append(xs[-1] + w)

    def grid():
        for x in xs:
            draw.line([(x, y0), (x, y0 + HEADER_H + ROW_H * (len(RECORDS) + 1))], fill=(60, 60, 60), width=3)
        yy = y0
        for _ in range(len(RECORDS) + 2):
            draw.line([(xs[0], yy), (xs[-1], yy)], fill=(60, 60, 60), width=3)
            yy += ROW_H if yy >= y0 + HEADER_H else HEADER_H

    grid()

    # 表头（印刷体）
    for j, h in enumerate(HEADERS):
        cx = (xs[j] + xs[j + 1]) / 2
        draw.text((cx, y0 + HEADER_H / 2), h, font=_font(FONT_SIZES["header"], False),
                  fill=(20, 20, 20), anchor="mm")

    # 数据行（楷体 + 随机抖动，模拟手写）
    y = y0 + HEADER_H
    for row in RECORDS:
        for j, val in enumerate(row):
            if not val:
                continue
            cx = (xs[j] + xs[j + 1]) / 2
            x = cx - draw.textlength(val, font=_font(FONT_SIZES["value"])) / 2
            for k, ch in enumerate(val):
                dx, dy = rng.uniform(-2.5, 2.5), rng.uniform(-3, 3)
                draw.text((x + dx, y + ROW_H / 2 - 15 + dy), ch,
                          font=_font(FONT_SIZES["value"]), fill=(25, 25, 90))
                x += draw.textlength(ch, font=_font(FONT_SIZES["value"]))
        y += ROW_H

    img.save(path)


def make_ground_truth(path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"file": os.path.basename(path).replace("ground_truth.json", "register_sheet_001.png"),
                   "headers": HEADERS, "records": RECORDS}, f, ensure_ascii=False, indent=2)


def make_ledger_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "用印台账"
    ws.append(HEADERS + ["转录状态", "复核标记"])
    for col, w in zip("ABCDEFGH", [6, 12, 44, 8, 12, 10, 12, 10]):
        ws.column_dimensions[col].width = w
    wb.save(path)


if __name__ == "__main__":
    os.makedirs(SHEET_DIR, exist_ok=True)
    img_path = os.path.join(SHEET_DIR, "register_sheet_001.png")
    make_sheet_image(img_path)
    make_ground_truth(os.path.join(SHEET_DIR, "ground_truth.json"))
    make_ledger_template(os.path.join(HERE, "sample", "ledger_template.xlsx"))
    print("已生成:", img_path)
