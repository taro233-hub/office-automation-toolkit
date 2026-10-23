# -*- coding: utf-8 -*-
"""供应商准入证照核验器
================================

解决的问题
----------
供应商准入（入库）时，要拿着准入登记表逐家核对营业执照 PDF：信用代码
格式与校验位、登记表与执照信息是否一致、营业期限是否已到期、经营范围能否
覆盖拟合作品类。人工逐项比对慢且容易看走眼，登记表填错一格后续合同、
开票全线返工。

核验项
------
1. 统一社会信用代码：18 位格式 + GB 32100-2015 校验位
2. 一致性：登记表 xlsx vs 执照 PDF 的 名称/信用代码/法定代表人/注册资本
3. 营业期限：是否已到期（对照系统当日）
4. 经营范围：登记的“推荐品类”是否原样出现在执照经营范围文本中
5. 输出：核验报告（md）+ 通过项的“准入填法文本块”（仍需人工复核）

用法
----
python verify.py                                   # 用 sample/ 下的虚构数据演示
python verify.py --intake 登记表.xlsx --licenses 执照目录
"""
import argparse
import datetime as dt
import os
import re

import pymupdf
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
WEIGHTS = [pow(3, i, 31) for i in range(17)]        # GB 32100-2015 权重 = 3^i mod 31

RX = {
    "统一社会信用代码": re.compile(r"统一社会信用代码[:：]?\s*([0-9A-Z]{18})"),
    "名称": re.compile(r"名\s*称[:：]?\s*([^\n]+)"),
    "法定代表人": re.compile(r"法定代表人[:：]?\s*([^\s，。；]+)"),
    "注册资本": re.compile(r"注册资本[:：]?\s*(\d+)\s*万"),
    "成立日期": re.compile(r"成立日期[:：]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日"),
    "营业期限": re.compile(r"营业期限[:：]?\s*自\d{4}年\d{1,2}月\d{1,2}日[起]?\s*至"
                          r"(?:(\d{4})年(\d{1,2})月(\d{1,2})日|长期|无固定期限)"),
    "经营范围": re.compile(r"经营范围[:：]?\s*([^\n]+)"),
}
INTAKE_COLS = ["供应商名称", "统一社会信用代码", "法定代表人", "注册资本(万元)",
               "推荐品类", "联系人", "联系电话", "营业执照文件"]


# ----------------------------------------------------------- 执照解析
_OCR = None


def _get_ocr():
    """PaddleOCR 懒加载单例（执照多为扫描件，无文字层时走 OCR）。"""
    global _OCR
    if _OCR is None:
        from paddleocr import PaddleOCR

        _OCR = PaddleOCR(lang="ch", use_doc_orientation_classify=False,
                         use_doc_unwarping=False, use_textline_orientation=False,
                         enable_mkldnn=False)   # 规避部分版本 oneDNN 的 PIR 算子兼容问题
    return _OCR


def parse_license(pdf_path):
    """从执照 PDF 提取文本并解析字段，返回 dict。
    优先取文字层；扫描件（无文字层）自动渲染成图后走 PaddleOCR。"""
    doc = pymupdf.open(pdf_path)
    text = "".join(page.get_text() for page in doc).strip()
    if len(text) < 20:                                  # 无文字层 → OCR
        parts = []
        for page in doc:
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2))
            png_path = pdf_path + ".tmp_ocr.png"
            pix.save(png_path)
            res = _get_ocr().predict(png_path)
            for r in res:
                parts.extend(str(t) for t in r["rec_texts"])
            os.remove(png_path)
        text = "\n".join(parts)
    doc.close()
    fields = {"_text": text}
    for key, rx in RX.items():
        m = rx.search(text)
        fields[key] = m.groups() if m else None
    return fields


def uscc_valid(code):
    """GB 32100-2015 校验位校验。"""
    if len(code) != 18 or any(c not in USCC_CHARS for c in code):
        return False, "格式不合法（需 18 位且不含 I/O/S/V/Z）"
    body, check = code[:17], code[17]
    s = sum(USCC_CHARS.index(c) * w for c, w in zip(body, WEIGHTS))
    expect = USCC_CHARS[(31 - s % 31) % 31]
    return check == expect, ("校验位应为 %s，实为 %s" % (expect, check)) if check != expect else ""


# ------------------------------------------------------------- 核验
def verify_one(row, lic, today):
    """对一家供应商执行全部核验，返回 (通过项, 问题项, 存疑项)。"""
    ok, bad, unsure = [], [], []

    def add(bucket, item):
        bucket.append(item)

    name = row["供应商名称"]
    lic_code = lic["统一社会信用代码"][0] if lic["统一社会信用代码"] else ""
    valid, why = uscc_valid(lic_code or row["统一社会信用代码"])
    if valid:
        add(ok, "信用代码校验位通过（GB 32100-2015）")
    else:
        add(bad, "信用代码校验位不通过：%s" % why)

    pairs = [("名称", row["供应商名称"], lic["名称"][0].strip() if lic["名称"] else ""),
             ("统一社会信用代码", row["统一社会信用代码"], lic_code),
             ("法定代表人", row["法定代表人"], lic["法定代表人"][0] if lic["法定代表人"] else "")]
    for field, a, b in pairs:
        if not b:
            add(unsure, "执照 PDF 未解析到「%s」，无法比对" % field)
        elif a != b:
            add(bad, "%s不一致：登记表 %r vs 执照 %r" % (field, a, b))
        else:
            add(ok, "%s一致（%s）" % (field, a))

    if lic["注册资本"]:
        if int(lic["注册资本"][0]) != int(row["注册资本(万元)"]):
            add(bad, "注册资本不一致：登记表 %s万 vs 执照 %s万" %
                (row["注册资本(万元)"], lic["注册资本"][0]))
        else:
            add(ok, "注册资本一致（%s万元）" % row["注册资本(万元)"])
    else:
        add(unsure, "执照未解析到注册资本（真实执照多为大写金额，需人工核对）")

    if lic["营业期限"] and lic["营业期限"][0]:
        y, m, d = map(int, lic["营业期限"][:3])
        if dt.date(y, m, d) < today:
            add(bad, "营业期限已于 %04d-%02d-%02d 到期" % (y, m, d))
        else:
            add(ok, "营业期限有效（至 %04d-%02d-%02d）" % (y, m, d))
    elif lic["营业期限"]:
        add(ok, "营业期限：长期")
    else:
        add(unsure, "执照未解析到营业期限")

    scope = lic["经营范围"][0] if lic["经营范围"] else ""
    cat = row["推荐品类"]
    if scope and cat and cat in scope:
        add(ok, "经营范围文本包含推荐品类「%s」（仍需人工判断业务资质）" % cat)
    elif scope:
        add(unsure, "经营范围未直接出现推荐品类「%s」，需人工判断：%s" % (cat, scope[:40]))
    else:
        add(unsure, "执照未解析到经营范围")

    return name, ok, bad, unsure


def make_fill_block(name, row, lic):
    """生成供人工复核的填法文本块（标注每个值的来源）。"""
    def src(field, value, lic_field):
        if lic[field]:
            return "%s（来源：执照PDF）" % value
        return "%s（来源：登记表，执照未解析到，待核）" % value

    lines = [
        "【%s】准入信息填法" % name,
        "供应商名称　：%s" % src("名称", row["供应商名称"], None),
        "统一社会信用代码：%s" % src("统一社会信用代码", row["统一社会信用代码"], None),
        "法定代表人　：%s" % src("法定代表人", row["法定代表人"], None),
        "注册资本　　：%s万元（来源：登记表，已与执照比对）" % row["注册资本(万元)"],
        "推荐品类　　：%s（来源：登记表）" % row["推荐品类"],
        "联系人/电话　：%s / %s（来源：登记表）" % (row["联系人"], row["联系电话"]),
    ]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="供应商准入证照核验")
    ap.add_argument("--intake", default=os.path.join(HERE, "sample", "supplier_intake.xlsx"))
    ap.add_argument("--licenses", default=os.path.join(HERE, "sample", "licenses"))
    ap.add_argument("--out", default=os.path.join(HERE, "output"))
    args = ap.parse_args()

    wb = load_workbook(args.intake, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    rows = [dict(zip(headers, [c.value for c in r])) for r in ws.iter_rows(min_row=2)]

    os.makedirs(args.out, exist_ok=True)
    blocks_dir = os.path.join(args.out, "fill_blocks")
    os.makedirs(blocks_dir, exist_ok=True)

    today = dt.date.today()
    report = ["# 供应商准入核验报告", "", "核验日期：%s" % today, ""]
    n_pass = n_bad = n_unsure = 0
    for row in rows:
        pdf = os.path.join(args.licenses, str(row["营业执照文件"]))
        if not os.path.exists(pdf):
            name = row["供应商名称"]
            print("[缺失] %s：找不到执照 %s" % (name, pdf))
            report.extend(["## %s" % name, "- △ 未找到执照 PDF：%s" % row["营业执照文件"],
                           "- **结论：暂缓，待补齐执照后核对**", ""])
            n_unsure += 1
            continue
        lic = parse_license(pdf)
        name, ok, bad, unsure = verify_one(row, lic, today)

        report.append("## %s" % name)
        for it in ok:
            report.append("- ✓ %s" % it)
        for it in bad:
            report.append("- ✗ %s" % it)
        for it in unsure:
            report.append("- △ %s" % it)
        if bad:
            report.append("- **结论：不通过，退回补正后再报**")
            n_bad += 1
        elif unsure:
            report.append("- **结论：暂缓，待人工核对 △ 项后录入**")
            n_unsure += 1
        else:
            report.append("- **结论：规则检查通过，人工核对原件与准入要求后再录入**")
            n_pass += 1
            with open(os.path.join(blocks_dir, "%s.txt" % name), "w", encoding="utf-8") as f:
                f.write(make_fill_block(name, row, lic))
        report.append("")

    report.insert(3, "**通过 %d 家 / 退回 %d 家 / 待人工核对 %d 家**" % (n_pass, n_bad, n_unsure))
    report_path = os.path.join(args.out, "verify_report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report))

    print("核验家数 : %d" % len(rows))
    print("结果     : 通过 %d / 退回 %d / 待人工核对 %d" % (n_pass, n_bad, n_unsure))
    print("报告     : %s" % report_path)
    print("填法文本 : %s" % blocks_dir)


if __name__ == "__main__":
    main()
