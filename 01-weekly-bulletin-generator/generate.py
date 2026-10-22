# -*- coding: utf-8 -*-
"""周进展通报一键生成器
================================

解决的问题
----------
每周一要把上周的《周进展通报》手工另存为新一周的版本：改标题和编制
日期、把上周“下周计划”搬进本周“本周进展”作为待填项、清空新的“下周
计划”、核对序号连续性。纯手工 15–30 分钟且容易漏行。

用法
----
# 由本周通报生成下周模板（周期自动 +7 天）
python generate.py --input sample/bulletin_2026-10-12.docx

# 指定下周起始日期
python generate.py --input sample/bulletin_2026-10-12.docx --date 2026-10-19

# 对某期通报做发布前校验（待填项 / 序号断档 / 空行）
python generate.py --input sample/bulletin_2026-10-19.docx --check
"""
import argparse
import datetime as dt
import os
import re
import shutil

from docx import Document

DATE_RE = re.compile(r"(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})")
COL_PROGRESS, COL_PLAN = "本周进展", "下周计划"   # 依赖的两个列名
PLACEHOLDER = "（待填）"
PLAN_PREFIX = "【承接上周计划】"


def parse_first_date(text):
    """从文本中解析第一个日期，返回 date 或 None。"""
    m = DATE_RE.search(text)
    if not m:
        return None
    y, mth, d = map(int, m.groups())
    try:
        return dt.date(y, mth, d)
    except ValueError:
        return None


def fmt(d):
    return f"{d.year}.{d.month:02d}.{d.day:02d}"


def week_range(monday):
    return f"{fmt(monday)}–{fmt(monday + dt.timedelta(days=4))}"


def find_bulletin_table(doc):
    """定位主表：表头同时含“本周进展”“下周计划”的表。"""
    for table in doc.tables:
        headers = [c.text.strip() for c in table.rows[0].cells]
        if COL_PROGRESS in headers and COL_PLAN in headers:
            return table, headers
    raise SystemExit("未找到包含“%s/%s”表头的表格" % (COL_PROGRESS, COL_PLAN))


def shift_week(text, old, new):
    """把文本里出现的周期/日期整体平移一周（兼容多种写法）。"""
    old_end = fmt(old + dt.timedelta(days=4))
    text = text.replace(f"{fmt(old)}–{old_end}", week_range(new))            # 2026.10.12–2026.10.16
    text = text.replace(f"{fmt(old)}–{old_end[5:]}", week_range(new))        # 2026.10.12–10.16
    text = text.replace(fmt(old), fmt(new))                                 # 2026.10.12
    return text.replace(old.isoformat(), new.isoformat())                   # 2026-10-12


def generate_next_week(src_path, next_monday=None):
    doc = Document(src_path)
    table, headers = find_bulletin_table(doc)
    ip, iplan = headers.index(COL_PROGRESS), headers.index(COL_PLAN)
    iserial = headers.index("序号") if "序号" in headers else None

    # 1) 解析当前周期，未指定则默认 +7 天
    cur = None
    for p in doc.paragraphs:
        cur = cur or parse_first_date(p.text)
    if cur is None:
        raise SystemExit("标题中未解析到日期，请用 --date 指定下周周一")
    nxt = next_monday or (cur + dt.timedelta(days=7))

    # 2) 标题/编制日期整体平移
    for p in doc.paragraphs:
        if parse_first_date(p.text):
            for run in p.runs:
                run.text = shift_week(run.text, cur, nxt)

    # 3) 行级搬运：上周“下周计划” → 本周“进展”的待填承接位
    carried, todo_progress, todo_plan = 0, 0, 0
    for n, row in enumerate(table.rows[1:], start=1):
        plan = row.cells[iplan].text.strip()
        progress_cell = row.cells[ip]
        if plan:
            progress_cell.text = f"{PLAN_PREFIX}{plan}\n{PLACEHOLDER}"
            carried += 1
        else:
            progress_cell.text = PLACEHOLDER
        todo_progress += 1
        row.cells[iplan].text = PLACEHOLDER
        todo_plan += 1
        if iserial is not None:
            row.cells[iserial].text = str(n)   # 序号重新对齐

    # 4) 输出（固定落在工具根目录 output/ 下）
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"bulletin_{nxt.isoformat()}.docx")
    doc.save(out_path)

    print("周期更新 : %s → %s" % (fmt(cur), fmt(nxt)))
    print("搬运条目 : %d 条（上周计划 → 本周进展承接位）" % carried)
    print("待填项   : 进展 %d 处 + 计划 %d 处" % (todo_progress, todo_plan))
    print("输出     :", out_path)
    return out_path


def check(src_path):
    """发布前校验：待填项、序号断档、整行空。"""
    doc = Document(src_path)
    table, headers = find_bulletin_table(doc)
    ip, iplan = headers.index(COL_PROGRESS), headers.index(COL_PLAN)
    iserial = headers.index("序号") if "序号" in headers else None

    issues = []
    serial_seen = []
    for n, row in enumerate(table.rows[1:], start=1):
        name = row.cells[headers.index("工作事项")].text.strip() if "工作事项" in headers else f"第{n}行"
        progress, plan = row.cells[ip].text.strip(), row.cells[iplan].text.strip()
        if PLACEHOLDER in progress or not progress:
            issues.append(f"[待填] {name} → 本周进展")
        if PLACEHOLDER in plan or not plan:
            issues.append(f"[待填] {name} → 下周计划")
        if not progress and not plan and not row.cells[1].text.strip():
            issues.append(f"[空行] 第{n}行整行为空")
        if iserial is not None:
            s = row.cells[iserial].text.strip()
            if s.isdigit():
                serial_seen.append(int(s))
            else:
                issues.append(f"[序号] 第{n}行序号缺失或不是数字: {s!r}")
    expect = list(range(1, len(table.rows)))
    if serial_seen and serial_seen != sorted(serial_seen):
        issues.append("[序号] 未按 1..N 顺序排列: %s" % serial_seen)
    elif serial_seen and serial_seen != expect:
        issues.append("[序号] 断档/重复: %s" % serial_seen)

    print("校验对象 :", src_path)
    print("命中问题 : %d 处" % len(issues))
    for it in issues:
        print("  -", it)
    if not issues:
        print("  ✓ 未发现问题，可发布")
    return issues


def main():
    ap = argparse.ArgumentParser(description="周进展通报一键生成/校验")
    ap.add_argument("--input", required=True, help="本周通报 docx 路径")
    ap.add_argument("--date", help="下周周一，格式 YYYY-MM-DD（默认自动 +7 天）")
    ap.add_argument("--check", action="store_true", help="只做发布前校验，不生成")
    args = ap.parse_args()

    if args.check:
        if check(args.input):
            raise SystemExit(1)
        return
    nxt = dt.date.fromisoformat(args.date) if args.date else None
    generate_next_week(args.input, nxt)


if __name__ == "__main__":
    main()
