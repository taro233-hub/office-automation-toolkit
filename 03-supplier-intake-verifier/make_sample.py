# -*- coding: utf-8 -*-
"""生成供应商准入核验的 demo 数据（全部虚构）：
1. sample/supplier_intake.xlsx —— 供应商准入登记表（4 家虚构供应商）
2. sample/licenses/*.pdf       —— 对应的“营业执照”PDF（仿版式、纯虚构）

其中故意埋了 3 处问题，用于演示 verify.py 的拦截能力：
- 锦岸幕墙：统一社会信用代码校验位错误
- 澜庭智能：营业期限已过期（2026-06-30 到期）
- 森启机电：登记表注册资本(500万) 与执照(300万) 不一致
"""
import os
import random

from PIL import Image, ImageDraw, ImageFont
import pymupdf
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
LICENSE_DIR = os.path.join(HERE, "sample", "licenses")

USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"      # GB 32100-2015 字符集
WEIGHTS = [pow(3, i, 31) for i in range(17)]        # 权重 = 3^i mod 31


def make_uscc(rng, broken=False):
    """生成 18 位统一社会信用代码；broken=True 时破坏校验位。"""
    code = "".join(rng.choice(USCC_CHARS) for _ in range(17))
    s = sum(USCC_CHARS.index(c) * w for c, w in zip(code, WEIGHTS))
    check = USCC_CHARS[(31 - s % 31) % 31]
    code += check
    if broken:                                        # 换掉校验位，保持格式合法
        alt = next(ch for ch in USCC_CHARS if ch != check)
        code = code[:-1] + alt
    return code


SUPPLIERS = [
    # 名称, 法定代表人, 注册资本(万元), 成立日期, 营业期限止, 经营范围, 登记表资本, 推荐品类, broken_uscc
    ("华曜建筑装饰工程有限公司", "陈志远", 1200, "2019年03月18日", "2039年03月17日",
     "建筑装饰工程、幕墙工程设计与施工；建筑材料销售。", 1200, "幕墙工程", False),
    ("锦岸幕墙科技有限公司", "林晓岚", 800, "2020年06月10日", "2040年06月09日",
     "建筑幕墙工程、金属门窗工程施工；幕墙材料研发与销售。", 800, "幕墙工程", True),
    ("澜庭智能科技有限公司", "周启铭", 500, "2018年11月02日", "2026年06月30日",
     "智能建筑系统、机电设备安装工程；软件开发。", 500, "机电设备", False),
    ("森启机电设备工程有限公司", "吴若晨", 300, "2021年09月15日", "2041年09月14日",
     "机电设备安装工程、暖通工程；机电设备销售。", 500, "机电设备", False),   # xlsx 填了 500，执照 300
]


def _font(size, bold=False):
    name = "simhei.ttf" if bold else "msyh.ttc"
    path = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", name)
    return ImageFont.truetype(path, size) if os.path.exists(path) else ImageFont.load_default()


def make_license_pdf(name, uscc, legal, capital_wan, founded, until, scope, path):
    """用 PIL 排版执照版式 → 插入 PDF 页面（demo 简化版式，非仿制真实证照）。"""
    W, H = 1000, 1414
    img = Image.new("RGB", (W, H), (252, 248, 236))
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, W - 20, H - 20], outline=(160, 30, 30), width=6)

    d.text((W / 2, 130), "营业执照（副本）", font=_font(56, True), fill=(120, 20, 20), anchor="mm")
    d.text((W / 2, 210), "【样张 · 虚构数据 · 仅供演示】", font=_font(26),
           fill=(150, 150, 150), anchor="mm")

    rows = [
        ("统一社会信用代码", uscc),
        ("名　　称", name),
        ("类　　型", "有限责任公司"),
        ("法定代表人", legal),
        ("注册资本", "%d万元人民币" % capital_wan),
        ("成立日期", founded),
        ("营业期限", "自%s起至%s" % (founded, until)),
        ("经营范围", scope),
        ("住　　所", "云溪市青澜区望江路88号"),
    ]
    y = 330
    for label, value in rows:
        d.text((90, y), label, font=_font(30), fill=(30, 30, 30))
        d.text((400, y), value, font=_font(30), fill=(30, 30, 30))
        y += 116
    d.text((W / 2, H - 90), "（示例生成件，无任何法律效力）", font=_font(24),
           fill=(150, 150, 150), anchor="mm")

    jpg = path.replace(".pdf", ".tmp.jpg")
    img.save(jpg, quality=88)                          # JPEG 嵌入，避免 PDF 存裸 RGB 导致单份 4MB
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)         # A4
    page.insert_image(page.rect, filename=jpg)
    doc.save(path)
    doc.close()
    os.remove(jpg)


def main():
    os.makedirs(LICENSE_DIR, exist_ok=True)
    rng = random.Random(7)
    codes = {}
    for s in SUPPLIERS:
        name, legal, cap, founded, until, scope, _sheet_cap, _cat, broken = s
        uscc = make_uscc(rng, broken=broken)
        codes[name] = uscc
        make_license_pdf(name, uscc, legal, cap, founded, until, scope,
                         os.path.join(LICENSE_DIR, "%s.pdf" % name))

    # 登记表里的信用代码必须与执照一致（含锦岸的坏校验位，才能演示“两处一致但都错”）
    xlsx_path = os.path.join(HERE, "sample", "supplier_intake.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "准入登记"
    ws.append(["供应商名称", "统一社会信用代码", "法定代表人", "注册资本(万元)",
               "推荐品类", "联系人", "联系电话", "营业执照文件"])
    for s in SUPPLIERS:
        name, legal, _cap, _f, _u, _sc, sheet_cap, cat, _b = s
        ws.append([name, codes[name], legal, sheet_cap, cat,
                   "X经理", "1380000%04d" % rng.randrange(1000, 9999),
                   "%s.pdf" % name])
    for col, w in zip("ABCDEFGH", [30, 22, 12, 14, 14, 10, 14, 30]):
        ws.column_dimensions[col].width = w
    wb.save(xlsx_path)
    print("已生成:", xlsx_path)
    print("执照目录:", LICENSE_DIR, "（%d 份）" % len(SUPPLIERS))


if __name__ == "__main__":
    main()
