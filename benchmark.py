"""拟真班级数据 + 评测：把「能不能做」变成数字。

造一份 52 人的名单和一堆真实风格的提交文件名（含老师实际会遇到的脏数据），
然后对比「朴素匹配」和「加了三道保险的匹配」谁更靠得住。
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from openpyxl import Workbook

import matcher as M

sys.stdout.reconfigure(encoding="utf-8")

HERE = Path(__file__).resolve().parent
DATA = HERE / "bench_data"
FOLDER = DATA / "作业文件夹"
ROSTER_XLSX = DATA / "全班名单.xlsx"

# --------------------------------------------------------------------------
# 名单：52 人
# --------------------------------------------------------------------------

ROSTER = [
    ("20210101", "张伟"), ("20210102", "张伟"), ("20210103", "陈峰"),
    ("20210104", "陈峰华"), ("20210105", "陈晓峰"), ("20210106", "李四"),
    ("20210107", "李四海"), ("20210108", "王五"), ("20210109", "赵六"),
    ("20210110", "欧阳娜"), ("20210111", "孙八"), ("20210112", "周杰"),
    ("20210113", "吴用"), ("20210114", "郑爽"), ("20210115", "王芳"),
    ("20210116", "冯磊"), ("20210117", "蒋雯"), ("20210118", "沈括"),
    ("20210119", "韩梅"), ("20210120", "杨帆"), ("20210121", "刘洋"),
    ("20210122", "黄磊"), ("20210123", "徐静"), ("20210124", "朱琳"),
    ("20210125", "高翔"), ("20210126", "林峰"), ("20210127", "何伟"),
    ("20210128", "罗静"), ("20210129", "邓超"), ("20210130", "曹阳"),
    ("20210131", "彭勇"), ("20210132", "曾毅"), ("20210133", "萧敬"),
    ("20210134", "田甜"), ("20210135", "于洋"), ("20210136", "余波"),
    ("20210137", "潘婷"), ("20210138", "秦岭"), ("20210139", "范冰"),
    ("20210140", "石磊"), ("20210141", "阎鹤"), ("20210142", "段誉"),
    ("20210143", "雷军"), ("20210144", "侯亮"), ("20210145", "白露"),
    ("20210146", "龙梅"), ("20210147", "万里"), ("20210148", "常青"),
    ("20210149", "邱枫"), ("20210150", "顾佳"), ("20210151", "武松"),
    ("20210152", "戴敏"),
]

# 事实：这 6 个人真的没交
UNSUBMITTED = {"20210102", "20210103", "20210106", "20210114", "20210119", "20210152"}

# --------------------------------------------------------------------------
# 文件夹里实际躺着的东西
# --------------------------------------------------------------------------

FILES = [
    # --- 学号命名，最稳 ---
    "20210101张伟-软件工程实验报告.docx",
    "杨帆-软工-20210120.docx",
    "吴用-计科2101-20260103.docx",
    # --- 重名陷阱：两个张伟，但 20210102 没交 ---
    "张伟.docx",
    # --- 子串陷阱：陈峰 / 陈峰华 / 陈晓峰 互相包含，而陈峰没交 ---
    "陈峰华_实验报告.docx",
    "陈晓峰.docx",
    # --- 子串陷阱：李四 / 李四海，而李四没交 ---
    "李四海 实验报告.pdf",
    # --- 各种真实命名风格 ---
    "20260103王五.zip",
    "赵六最终版.docx",
    "第三组_欧阳娜_实验报告.pdf",
    "孙八 周杰 小组作业.docx",
    "王芳 作业.docx",
    "冯磊.zip",
    "蒋雯/实验一.docx",
    "蒋雯/实验二.docx",
    "沈括.doc",
    "黄磊_作业_终稿_真的终稿.docx",
    "徐静(1).docx",
    "徐静(2).docx",
    "朱琳实验报告.pdf",
    "高翔-2026-实验.docx",
    "林峰.docx",
    "刘洋.docx",
    "何伟作业.docx",
    "罗静.docx",
    "邓超_实验报告.docx",
    "曹阳.zip",
    "彭勇-作业.pdf",
    "曾毅.docx",
    "萧敬_报告.docx",
    "田甜 作业.docx",
    "于洋.docx",
    "余波.docx",
    "潘婷实验.docx",
    "秦岭_实验一.docx",
    "范冰.docx",
    "石磊 实验报告.docx",
    "阎鹤.docx",
    "段誉-作业.docx",
    "雷军.docx",
    "侯亮作业.docx",
    "白露.docx",
    "龙梅_实验.docx",
    "万里.docx",
    "常青 报告.docx",
    "邱枫.docx",
    "顾佳作业.docx",
    "武松.docx",
    # --- 噪声：韩梅没交，只是打开过 Word；别把她算成交了 ---
    "~$韩梅_报告终稿.docx",
    "Thumbs.db",
    ".DS_Store",
    # --- 认不出是谁的 ---
    "微信图片_20260103120000.jpg",
    "全班作业汇总.zip",
]


def build() -> None:
    if DATA.exists():
        shutil.rmtree(DATA)
    FOLDER.mkdir(parents=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "名单"
    ws.append(["学号", "姓名"])
    for sid, name in ROSTER:
        ws.append([sid, name])
    wb.save(ROSTER_XLSX)

    for rel in FILES:
        p = FOLDER / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("x", encoding="utf-8")


# --------------------------------------------------------------------------
# 朴素对照：不过滤噪声、姓名随便子串命中、重名随便认一个
# --------------------------------------------------------------------------


def naive_match(roster, root: Path):
    """真正「随手写」的版本：每个人的名字分别去文件名里找子串，不遮蔽、
    不排长度、不过滤临时文件。这是大多数人第一版会写出来的东西。"""
    everything = sorted(p for p in root.rglob("*") if p.is_file())
    submitted = set()
    by_name = {M.normalize(p.name): p for p in roster}
    for f in everything:
        text = M.normalize(f.stem)
        for n, p in by_name.items():
            if n in text:
                submitted.add(p.key)
        for sid in M.extract_ids(text):
            for p in roster:
                if p.sid and M.normalize(p.sid) == sid:
                    submitted.add(p.key)
                    break
    return submitted


# --------------------------------------------------------------------------
# 跑
# --------------------------------------------------------------------------


def evaluate(name, submitted, unsure, missing, unidentified, total_unsure):
    truth_submitted = {sid for sid, _ in ROSTER} - UNSUBMITTED
    truth_missing = UNSUBMITTED

    fp = submitted - truth_submitted  # 说交了，其实没交  ← 老师漏收
    fn = {p.key for p in missing} & truth_submitted  # 说没交，其实交了 ← 冤枉人

    print(f"\n{'=' * 62}\n{name}\n{'=' * 62}")
    print(f"  已交（确定）  {len(submitted):>3} 人")
    print(f"  待确认        {total_unsure:>3} 人")
    print(f"  未交          {len(missing):>3} 人")
    print(f"  认不出的文件  {len(unidentified):>3} 个")
    print(f"  ---")
    print(f"  误判为已交（危险，会漏收作业）  {len(fp):>3}")
    if fp:
        names = {sid: n for sid, n in ROSTER}
        print(f"      {[f'{names[s]}({s})' for s in sorted(fp)]}")
    print(f"  误判为未交（冤枉学生）          {len(fn):>3}")
    if fn:
        names = {sid: n for sid, n in ROSTER}
        print(f"      {[f'{names[s]}({s})' for s in sorted(fn)]}")
    exact = len(submitted & truth_submitted)
    print(f"  正确识别已交                    {exact:>3} / {len(truth_submitted)}")
    return {"fp": fp, "fn": fn, "exact": exact}


def main() -> None:
    build()
    roster = M.read_roster(ROSTER_XLSX)
    files = M.collect_files(FOLDER)
    print(f"名单读入 {len(roster)} 人；文件夹里 {len(FILES)} 个路径，"
          f"过滤掉噪声后参与匹配 {len(files)} 个")

    # --- 对照：朴素做法 ---
    naive = naive_match(roster, FOLDER)
    truth = {sid for sid, _ in ROSTER} - UNSUBMITTED
    naive_fp = naive - truth
    naive_fn = truth - naive
    print(f"\n{'=' * 62}\n朴素匹配（不过滤临时文件 / 姓名任意子串 / 重名随手认）\n{'=' * 62}")
    print(f"  已交  {len(naive):>3} 人")
    print(f"  误判为已交（危险）  {len(naive_fp):>3}   "
          f"{[f'{n}({s})' for s, n in ROSTER if s in naive_fp]}")
    print(f"  误判为未交         {len(naive_fn):>3}   "
          f"{[f'{n}({s})' for s, n in ROSTER if s in naive_fn]}")

    # --- 正式：三道保险 ---
    res = M.match_files(roster, files)
    stats = evaluate(
        "本方案（学号优先 + 最长优先遮蔽 + 噪声过滤 + 三态输出）",
        set(res.submitted), res.unsure, res.missing, res.unidentified,
        len(res.unsure),
    )

    print(f"\n{'=' * 62}\n给老师看的输出\n{'=' * 62}")
    print(f"  ✓ 已交 {len(res.submitted)}")
    print(f"  ? 待确认 {len(res.unsure)}")
    print(f"  ✕ 未交 {len(res.missing)}")
    if res.unsure:
        print("\n  待确认（需要人工扫一眼）：")
        for key, hits in res.unsure.items():
            who = hits[0].person
            print(f"     {who}  ← 疑似来自 {[h.path.name for h in hits]}")
    if res.missing:
        print("\n  未交名单：")
        print("     " + "、".join(str(p) for p in res.missing))
    if res.unidentified:
        print("\n  认不出是谁的文件（可能有漏网之交）：")
        for f in res.unidentified:
            print(f"     {f.relative_to(FOLDER)}")

    # --- 可解释性：每个判定都能说出证据 ---
    print(f"\n{'=' * 62}\n判定依据（老师要能核对，才敢拿去催人）\n{'=' * 62}")
    for key in list(res.submitted)[:6]:
        for h in res.submitted[key]:
            print(f"  {h.person}  ←  {h.path.name}   [{h.how}]")


if __name__ == "__main__":
    main()
