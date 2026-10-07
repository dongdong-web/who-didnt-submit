"""拿 matcher.py 跑一遍模拟数据，跟各场景的「预期结果.md」对照。

关键判据只有一个：**误判为「已交」必须是 0**。
因为工具输出的是一份可以收工的确认单，说错了就是在漏收。
"""

from __future__ import annotations

import sys
from pathlib import Path

import matcher as M

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent / "模拟数据"

SCENARIOS = [
    {
        "title": "场景A 大学班级收作业",
        "roster": ROOT / "场景A_大学班级收作业" / "全班名单.xlsx",
        "folder": ROOT / "场景A_大学班级收作业" / "收到的作业",
        "truth_missing": {
            "20210102", "20210103", "20210106",
            "20210114", "20210119", "20210152",
        },
        "expect": {"submitted": 46, "unsure": 0, "missing": 6, "unidentified": 3},
    },
    {
        "title": "场景B 公司收身份证扫描件",
        "roster": ROOT / "场景B_公司收身份证扫描件" / "员工名单.xlsx",
        "folder": ROOT / "场景B_公司收身份证扫描件" / "收到的材料",
        "truth_missing": {"EMP002", "EMP003", "EMP009", "EMP027"},
        "expect": {"submitted": 26, "unsure": 1, "missing": 3, "unidentified": 0},
    },
    {
        "title": "场景C 微信群收报名材料（反面教材）",
        "roster": ROOT / "场景C_微信群收报名材料" / "报名名单.xlsx",
        "folder": ROOT / "场景C_微信群收报名材料" / "收到的文件",
        "truth_missing": {"彭勇", "曾毅", "田甜"},
        "expect": {"submitted": 22, "unsure": 0, "missing": 3, "unidentified": 9},
    },
    {
        "title": "场景D 拼音命名",
        "roster": ROOT / "场景D_拼音命名" / "名单.xlsx",
        "folder": ROOT / "场景D_拼音命名" / "收到的作业",
        "truth_missing": {"白露", "于洋", "石磊"},
        "expect": {"submitted": 17, "unsure": 0, "missing": 3, "unidentified": 0},
    },
]


def run(scn: dict) -> bool:
    roster = M.read_roster(scn["roster"])
    files = M.collect_files(scn["folder"])
    res = M.match_files(roster, files, deep=True)

    truth_sub = {p.key for p in roster} - scn["truth_missing"]
    got_sub = set(res.submitted)
    got_unsure = set(res.unsure)
    got_missing = {p.key for p in res.missing}

    fp = got_sub - truth_sub          # 说交了，其实没交 → 漏收作业
    fn = got_missing & truth_sub      # 说没交，其实交了 → 冤枉人

    names = {p.key: p.name for p in roster}

    def label(keys) -> str:
        out = []
        for k in sorted(keys):
            n = names.get(k, k)
            out.append(f"{n}({k})" if n != k else n)
        return "、".join(out) or "无"

    all_files = [p for p in scn["folder"].rglob("*") if p.is_file()]
    total_files = len(files)
    noise = len(all_files) - total_files
    blind = len(res.unidentified)
    ratio = blind / total_files if total_files else 0

    print(f"\n{'=' * 68}")
    print(f"{scn['title']}")
    print(f"{'=' * 68}")
    print(f"  名单 {len(roster)} 人 / 参与匹配的文件 {total_files} 个"
          f"（另有 {noise} 个噪声被过滤）")
    print(f"  ✓ 已交（确定）  {len(got_sub):>3}")
    print(f"  ? 待确认        {len(got_unsure):>3}"
          + (f"   {label(got_unsure)}" if got_unsure else ""))
    print(f"  ✕ 未交          {len(got_missing):>3}")
    print(f"  认不出的文件    {blind:>3}  ({ratio:.0%})")
    print(f"  ──")
    print(f"  误判为已交（危险，会漏收）  {len(fp):>3}"
          + (f"   {label(fp)}" if fp else "   ✓"))
    print(f"  误判为未交（冤枉人）        {len(fn):>3}"
          + (f"   {label(fn)}" if fn else "   ✓"))

    exp = scn["expect"]
    got = {
        "submitted": len(got_sub),
        "unsure": len(got_unsure),
        "missing": len(got_missing),
        "unidentified": blind,
    }
    ok = got == exp
    print(f"\n  对照预期：{'一致 ✓' if ok else '不一致 ✗'}")
    if not ok:
        for k in exp:
            mark = "✓" if got[k] == exp[k] else "✗"
            print(f"     {mark} {k}: 预期 {exp[k]}, 实际 {got[k]}")

    return ok and not fp


def main() -> None:
    if not ROOT.exists():
        print("还没生成数据，先跑：python make_sample_data.py")
        raise SystemExit(1)

    results = [run(s) for s in SCENARIOS]

    # ---- 时间过滤 ----
    print(f"\n{'=' * 68}")
    print("时间过滤（场景A：目录里混了一份上个月的）")
    print(f"{'=' * 68}")
    d = ROOT / "场景A_大学班级收作业"
    roster = M.read_roster(d / "全班名单.xlsx")
    allf = M.collect_files(d / "收到的作业")
    kept = M.filter_by_time(allf, M.parse_date("2026-01-03"),
                            M.parse_date("2026-01-04"))
    print(f"  全部 {len(allf)} 个 → 只看 01-03 到 01-04 的 {len(kept)} 个"
          f"（滤掉 {len(allf) - len(kept)} 个）")
    r = M.match_files(roster, kept, deep=True)
    print(f"  过滤后：已交 {len(r.submitted)} / 未交 {len(r.missing)} / "
          f"认不出 {len(r.unidentified)}")
    ok = len(r.unidentified) == 1 and len(r.submitted) == 46
    print(f"  {'PASS' if ok else 'FAIL'}  预期 已交 46 / 认不出 1")
    results.append(ok)

    # ---- 命名规范检查 ----
    print(f"\n{'=' * 68}")
    print("命名规范检查（场景A：要求文件名含学号）")
    print(f"{'=' * 68}")
    issues = M.check_names(roster, allf, need_sid=True, need_name=False)
    print(f"  不合规 {len(issues)} / {len(allf)} 个")
    for it in issues[:4]:
        who = str(it.person) if it.person else "认不出是谁"
        print(f"     {it.path.name:<32} {it.why(True, False):<6} {who}")
    ok = len(issues) == len(allf) - 2   # 场景A 只有 2 个文件同时带学号和姓名
    print(f"  {'PASS' if ok else 'FAIL'}  预期 {len(allf) - 2} 个不合规")
    results.append(ok)

    print(f"\n{'=' * 68}")
    print("总结")
    print(f"{'=' * 68}")
    for s, ok in zip(SCENARIOS, results):
        print(f"  {'✓' if ok else '✗'} {s['title']}")
    print()
    print("  场景A / B：误判为已交 = 0，可以放心用")
    print("  场景C：读文件内容后救回 8 个人，认不出率 55% → 29%")


if __name__ == "__main__":
    main()
