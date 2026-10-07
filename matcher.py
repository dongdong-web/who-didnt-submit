"""「谁没交？」—— 文件夹名单核对器，匹配核心（原型 v1）

三条设计原则：

  1. 学号优先于姓名。学号唯一，姓名会重名。
  2. 姓名匹配「最长优先 + 遮蔽」。名单里同时有 陈峰 / 陈峰华 / 陈晓峰 时，
     `陈峰华_实验报告.docx` 只能算陈峰华，绝不能顺手把陈峰也算成交了。
  3. 拿不准就说拿不准。输出三态（已交 / 待确认 / 未交）。
     老师漏收一份作业的代价，远大于多问一句。
"""

from __future__ import annotations

import itertools
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

# --------------------------------------------------------------------------
# 归一化
# --------------------------------------------------------------------------

_INVISIBLE = re.compile(r"[\s\u00a0\u200b-\u200f\u2060\ufeff]+")
_ID_PATTERNS = (
    re.compile(r"\d{6,}"),            # 纯数字学号：20210101
    re.compile(r"[a-z]{1,6}\d{2,}"),  # 带字母前缀的工号：emp001 / no12
)
_NOISE = {"thumbs.db", "desktop.ini", ".ds_store", "icon\r"}

_MASK = "\x00"


def normalize(text: str) -> str:
    """全角转半角、转小写、去掉空白与零宽字符。"""
    if not text:
        return ""
    return _INVISIBLE.sub("", unicodedata.normalize("NFKC", text).lower())


def is_noise(path: Path) -> bool:
    """系统文件和 Office 临时文件。漏掉它们会直接造成误判。

    `~$韩梅.docx` 是韩梅打开过 Word 留下的锁文件，不代表她交了作业。
    """
    name = path.name.lower()
    return name in _NOISE or name.startswith("~$") or name.startswith(".")


def extract_ids(text: str) -> list[str]:
    """从归一化后的文本里挖出所有像学号 / 工号的片段。"""
    out: list[str] = []
    for pat in _ID_PATTERNS:
        out.extend(pat.findall(text))
    return out


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Person:
    name: str
    sid: str = ""

    @property
    def key(self) -> str:
        return self.sid or self.name

    def __str__(self) -> str:
        return f"{self.name}({self.sid})" if self.sid else self.name


@dataclass
class Hit:
    person: Person
    path: Path
    how: str  # "学号" | "姓名"
    sure: bool  # True = 算已交；False = 只能待确认


@dataclass
class Result:
    submitted: dict[str, list[Hit]]  # 确定已交
    unsure: dict[str, list[Hit]]  # 待确认
    missing: list[Person]  # 未交
    unidentified: list[Path]  # 认不出是谁的文件

    @property
    def submitted_count(self) -> int:
        return len(self.submitted)


# --------------------------------------------------------------------------
# 读名单
# --------------------------------------------------------------------------

_NAME_KEYS = ("姓名", "名字", "学生", "成员", "人员", "同学")
_SID_KEYS = ("学号", "工号", "编号", "序号", "id", "账号")


def read_roster(path: Path) -> list[Person]:
    """从 Excel 读名单。容忍有表头 / 没表头 / 多列。"""
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = [
        list(r)
        for r in ws.iter_rows(values_only=True)
        if any(c is not None and str(c).strip() for c in r)
    ]
    wb.close()
    if not rows:
        return []

    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    lowered = [h.lower() for h in header]

    name_col = sid_col = None
    for i, h in enumerate(lowered):
        if name_col is None and any(k in h for k in _NAME_KEYS):
            name_col = i
        if sid_col is None and any(k in h for k in _SID_KEYS):
            sid_col = i

    if name_col is None and sid_col is None:
        # 没表头：靠内容猜。整列长数字 → 学号；中文 2~4 字 → 姓名。
        data = rows
        body = [r for r in data[:50]]
        for i in range(max(len(r) for r in body)):
            col = [str(r[i]).strip() for r in body if i < len(r) and r[i] is not None]
            if not col:
                continue
            if sid_col is None and all(re.fullmatch(r"\d{6,}", c) for c in col):
                sid_col = i
            elif name_col is None and all(
                re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", c) for c in col
            ):
                name_col = i
    else:
        data = rows[1:]

    if name_col is None:
        return []

    people: list[Person] = []
    seen: set[str] = set()
    for r in data:
        if name_col >= len(r) or r[name_col] is None:
            continue
        name = str(r[name_col]).strip()
        if not name:
            continue
        sid = ""
        if sid_col is not None and sid_col < len(r) and r[sid_col] is not None:
            sid = str(r[sid_col]).strip()
            sid = re.sub(r"\.0$", "", sid)
        key = sid or name
        if key in seen:
            continue
        seen.add(key)
        people.append(Person(name, sid))
    return people


# --------------------------------------------------------------------------
# 匹配
# --------------------------------------------------------------------------


def collect_files(root: Path, recursive: bool = True) -> list[Path]:
    it = root.rglob("*") if recursive else root.glob("*")
    return sorted(p for p in it if p.is_file() and not is_noise(p))


_LABEL_RE = re.compile(r"(?:姓\s*名|名\s*字|学\s*号|工\s*号|编\s*号)")


def signature_zone(text: str, window: int = 14) -> str:
    """把「姓名：X」「学号：Y」这类标签后面的一小段截出来拼在一起。

    只在这一小段里找名字，比在全文里找安全得多 ——
    正文里提到别人（"参考了李四同学的代码"）不会误伤。
    """
    chunks = []
    for m in _LABEL_RE.finditer(text):
        chunks.append(text[m.start():m.end() + window])
    return "\n".join(chunks)


def pinyin_variants(name: str, min_len: int = 4) -> set[str]:
    """把一个中文名转成所有可能的拼音写法（含多音字）。

    必须用多音字全变体：「曾毅」的「曾」作姓读 zēng，但 pypinyin 默认给 céng，
    只取默认读音的话 `zengyi.docx` 就永远匹配不上。

    `min_len=4` 会挡掉太短的拼音：单字名（「李」→ li）会撞上大量英文单词，
    收进来全是误报。
    """
    try:
        from pypinyin import Style, pinyin as _pinyin
    except Exception:
        return set()
    try:
        parts = _pinyin(name, style=Style.NORMAL, heteronym=True, errors="ignore")
    except Exception:
        return set()
    if not parts:
        return set()
    out: set[str] = set()
    for combo in itertools.product(*parts):
        s = "".join(combo)
        if len(s) >= min_len:
            out.add(s)
    return out


class _Index:
    """把名单预处理成几张查表，省得每匹配一个文件就重算一遍。"""

    def __init__(self, roster: list[Person]):
        self.by_sid: dict[str, Person] = {
            normalize(p.sid): p for p in roster if p.sid
        }

        self.by_name: dict[str, list[Person]] = {}
        for p in roster:
            self.by_name.setdefault(normalize(p.name), []).append(p)
        self.names_desc = sorted(self.by_name, key=len, reverse=True)

        self.by_pinyin: dict[str, list[Person]] = {}
        for p in roster:
            for py in pinyin_variants(p.name):
                self.by_pinyin.setdefault(py, []).append(p)
        self.py_desc = sorted(self.by_pinyin, key=len, reverse=True)


def _find_people(idx: _Index, text: str, strong: bool, label: str):
    """在一段文本里找名单上的人，返回 [(Person, 怎么认出来的, 是否确定)]。

    strong=True 表示这段文本是文件名或署名区，里面的名字可信；
    strong=False 表示只是正文里提到，只能算「待确认」。
    """
    out: list[tuple[Person, str, bool]] = []

    # 编号优先：学号 / 工号唯一，能直接定人 —— 重名也救得回来
    for sid in extract_ids(text):
        if sid in idx.by_sid:
            out.append((idx.by_sid[sid], f"{label}·编号", True))
            return out

    shadow = text

    # 中文姓名：最长优先 + 遮蔽
    for n in idx.names_desc:
        if n and n in shadow:
            # 遮蔽：占掉的位置不能再被更短的名字命中
            shadow = shadow.replace(n, _MASK * len(n))
            group = idx.by_name[n]
            for person in group:
                sure = strong and len(group) == 1
                out.append((person, f"{label}·{'署名' if strong else '提及'}", sure))

    # 拼音：chenfeng.docx / zengyi.docx
    for py in idx.py_desc:
        if py in shadow:
            shadow = shadow.replace(py, _MASK * len(py))
            group = idx.by_pinyin[py]
            for person in group:
                sure = strong and len(group) == 1
                out.append((person, f"{label}·拼音", sure))

    return out


def match_files(
    roster: list[Person],
    files: list[Path],
    use_parent: bool = True,
    deep: bool = False,
) -> Result:
    """deep=True 时，文件名认不出（或拿不准）的文件会再读一遍内容找署名。

    分两阶段是有意的：先过文件名，只有没定下来的才去读内容。
    读文件是慢操作，能少读一个是一个。
    """
    idx = _Index(roster)

    resolved: dict[Path, list[Hit]] = {}
    pending: list[Path] = []
    unidentified: list[Path] = []

    # ---- 阶段一：文件名（含父目录名） ----
    for f in files:
        parts = [f.stem]
        if use_parent and f.parent.name:
            # 有人用「陈峰/实验一.docx」这种个人文件夹交作业
            parts.append(f.parent.name)
        text = normalize(" ".join(parts))

        found = _find_people(idx, text, True, "文件名")
        if found:
            hits = [Hit(p, f, how, sure) for p, how, sure in found]
            resolved[f] = hits
            if not any(h.sure for h in hits):
                pending.append(f)      # 只是「待确认」，值得再读内容定死
        else:
            pending.append(f)

    # ---- 阶段二：文件内容 ----
    extract_text = None
    if deep:
        try:
            from content import extract_text as _extract
            extract_text = _extract
        except Exception:
            extract_text = None

    if extract_text is not None and pending:
        still: list[Path] = []
        for f in pending:
            body = extract_text(f)
            found: list = []
            if body:
                zone = signature_zone(body)
                if zone:
                    found = _find_people(idx, normalize(zone), True, "内容")
                if not found:
                    # 全文里提到名字：弱证据，只敢标「待确认」
                    found = _find_people(idx, normalize(body), False, "内容")
            if found:
                resolved[f] = [Hit(p, f, how, sure) for p, how, sure in found]
            elif f not in resolved:
                still.append(f)
        unidentified = still
    else:
        unidentified = [f for f in pending if f not in resolved]

    # ---- 汇总 ----
    submitted: dict[str, list[Hit]] = {}
    unsure: dict[str, list[Hit]] = {}
    for hits in resolved.values():
        for hit in hits:
            bucket = submitted if hit.sure else unsure
            bucket.setdefault(hit.person.key, []).append(hit)

    # 确定命中的，从待确认里摘掉
    for key in list(unsure):
        if key in submitted:
            del unsure[key]

    hit_keys = set(submitted) | set(unsure)
    missing = [p for p in roster if p.key not in hit_keys]

    return Result(submitted, unsure, missing, unidentified)


# --------------------------------------------------------------------------
# 文件类型：微信的文件目录里混着表情包、随手发的图、视频
# --------------------------------------------------------------------------

# key -> (界面上显示的名字, 扩展名集合)。other 是兜底，扩展名留空。
#
# 名字用老师嘴里的说法（Word / PPT / PDF），而不是「文档 / 演示」这种
# 按格式划分的行话 —— 后者得先在脑子里翻译一遍。
FILE_KINDS: dict[str, tuple[str, set[str]]] = {
    "word": ("Word", {".doc", ".docx", ".rtf", ".odt", ".wps"}),
    "excel": ("Excel", {".xls", ".xlsx", ".xlsm", ".csv", ".ods", ".et"}),
    "ppt": ("PPT", {".ppt", ".pptx", ".odp", ".dps"}),
    "pdf": ("PDF", {".pdf"}),
    "zip": ("压缩包", {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"}),
    "image": ("图片", {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
                       ".heic", ".tif", ".tiff", ".svg"}),
    "video": ("视频", {".mp4", ".mov", ".avi", ".mkv", ".flv", ".wmv",
                       ".m4v", ".webm"}),
    "other": ("其他", set()),
}

# 界面上按这个顺序排
KIND_ORDER = ("word", "excel", "ppt", "pdf", "zip", "image", "video", "other")

# 默认一个都不勾 = 不限制。勾一个就是「只看这个」——
# 这样默认状态是干净的，也不会替用户做有风险的假设。
DEFAULT_KINDS: set[str] = set()

_EXT_TO_KIND = {
    ext: key for key, (_label, exts) in FILE_KINDS.items() for ext in exts
}


def kind_of(path: Path) -> str:
    """这个文件属于哪一类。认不出来的都归 other。"""
    return _EXT_TO_KIND.get(path.suffix.lower(), "other")


def filter_by_kind(files, kinds) -> list[Path]:
    """只留下指定类型的文件。kinds 是 key 的集合。"""
    return [f for f in files if kind_of(f) in kinds]


def kind_counts(files) -> dict[str, int]:
    """数一数文件夹里各类文件各有多少 —— 好让用户知道勾哪些有意义。"""
    out = {k: 0 for k in KIND_ORDER}
    for f in files:
        out[kind_of(f)] += 1
    return out


# --------------------------------------------------------------------------
# 时间过滤：微信的 FileStorage 按月份堆着所有会话的文件
# --------------------------------------------------------------------------


def parse_date(text: str) -> datetime | None:
    """认几种常见写法：2026-01-03 / 20260103 / 2026-01 / 2026。"""
    s = (text or "").strip().replace("/", "-").replace(".", "-")
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y-%m", "%Y%m", "%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def filter_by_time(files, since=None, until=None):
    """按文件修改时间筛。

    Windows 复制和另存为都会保留原修改时间，所以对「从微信目录拷出来」
    的文件照样管用。
    """
    out = []
    for f in files:
        try:
            mtime = f.stat().st_mtime
        except OSError:
            continue
        if since is not None and mtime < since.timestamp():
            continue
        if until is not None and mtime > until.timestamp() + 86399:  # 含当天
            continue
        out.append(f)
    return out


# --------------------------------------------------------------------------
# 微信目录：存储位置可以改，而且根目录里塞满了图片视频
# --------------------------------------------------------------------------


def locate_file_dir(root: Path, max_depth: int = 3, max_dirs: int = 800,
                    max_seconds: float = 1.5) -> Path:
    """如果拖进来的是微信的存储根，自动钻到真正放文件的那一层。

    微信的存储位置能在「设置 → 文件管理」里改（比如 D:\\mybin\\xwechat_files），
    所以路径不能写死。而根目录里塞着图片、视频、消息数据库，动辄几万个文件 ——
    直接扫又慢又没意义，我们只关心别人发来的文档。

    **三重限流，缺一不可**（拖 C:\\Windows 进来实测过）：
      - max_dirs：最多翻这么多个目录
      - max_entries：单个目录超过这么多子目录就放弃 —— WinSxS 那种
        一次 iterdir 就是几十万条，只数目录个数根本挡不住（实测 30 秒）
      - max_seconds：总耗时兜底

    找得到就返回那一层，找不到原样返回。
    """
    import time

    deadline = time.monotonic() + max_seconds
    budget = [max_dirs]

    def walk(d: Path, depth: int) -> Path | None:
        if depth > max_depth or budget[0] <= 0 or time.monotonic() > deadline:
            return None
        budget[0] -= 1
        try:
            subs: dict[str, Path] = {}
            for p in d.iterdir():
                if len(subs) > 3000:      # 这目录大得离谱，别翻了
                    return None
                try:
                    if p.is_dir():
                        subs[p.name.lower()] = p
                except OSError:
                    continue
        except OSError:
            return None

        # 微信 3.x 是 <账号>/FileStorage/File，4.x 是 <账号>/msg/file
        for holder_name, target in (("filestorage", "file"), ("msg", "file")):
            holder = subs.get(holder_name)
            if holder is None:
                continue
            try:
                for p in holder.iterdir():
                    if p.is_dir() and p.name.lower() == target:
                        return p
            except OSError:
                pass

        for sub in subs.values():
            got = walk(sub, depth + 1)
            if got:
                return got
        return None

    if not root.is_dir():
        return root
    return walk(root, 1) or root


def quick_scan(root: Path, limit: int = 3000):
    """轻量扫一眼：最多取 limit 个文件，边遍历边停，不排序也不全量收集。

    拖放之后要立刻给用户反馈，不能拿 collect_files 那种全量遍历去堵界面。
    返回 (文件列表, 是否被截断)。
    """
    out = []
    try:
        for p in root.rglob("*"):
            if len(out) >= limit:
                return out, True
            try:
                if p.is_file() and not is_noise(p):
                    out.append(p)
            except OSError:
                continue
    except OSError:
        pass
    return out, False


# --------------------------------------------------------------------------
# 命名规范检查：只看文件名，不看内容
# --------------------------------------------------------------------------


@dataclass
class NameIssue:
    path: Path
    has_sid: bool
    has_name: bool
    person: Person | None
    order_bad: bool = False

    def why(self, need_sid: bool, need_name: bool) -> str:
        miss = []
        if need_sid and not self.has_sid:
            miss.append("缺学号")
        if need_name and not self.has_name:
            miss.append("缺姓名")
        if self.order_bad:
            miss.append("顺序不对")
        return "、".join(miss) or "—"


def _name_positions(idx: _Index, text: str) -> tuple[int, int]:
    """文件名里 学号 / 姓名的起始位置，-1 表示没有。"""
    sid_pos = -1
    for s in extract_ids(text):
        if s in idx.by_sid:
            sid_pos = text.find(s)
            break

    name_pos = -1
    for n in idx.names_desc:
        if n and n in text:
            name_pos = text.find(n)
            break
    if name_pos < 0:
        for py in idx.py_desc:
            p = text.find(py)
            if p >= 0:
                name_pos = p
                break
    return sid_pos, name_pos


def name_stats(roster, files) -> dict:
    """先看看大家实际都怎么命名的 —— 好帮你决定要求该定多严。

    这个比「自动生成正则」有用：规范是你定的，但定之前得知道
    「要求写学号的话，会牵连多少人」。
    """
    idx = _Index(roster)
    has_sid = has_name = has_both = sid_first = name_first = 0
    for f in files:
        text = normalize(f"{f.stem} {f.parent.name}")
        sid_pos, name_pos = _name_positions(idx, text)
        if sid_pos >= 0:
            has_sid += 1
        if name_pos >= 0:
            has_name += 1
        if sid_pos >= 0 and name_pos >= 0:
            has_both += 1
            if sid_pos < name_pos:
                sid_first += 1
            else:
                name_first += 1
    return {
        "total": len(files),
        "has_sid": has_sid,
        "has_name": has_name,
        "has_both": has_both,
        "sid_first": sid_first,
        "name_first": name_first,
    }


def check_names(roster, files, need_sid=True, need_name=True, order=None):
    """挑出文件名不合规的。

    order 可以是 None（不查顺序）/ "sid_first"（学号在前）/ "name_first"。

    这里**只看文件名，不看内容** —— 老师要求的是「文件名」规范，
    内容里写得再全也不算数。这跟 match_files 是两回事：
    文件名不规范的文件，照样可能被匹配成功。
    """
    idx = _Index(roster)
    issues: list[NameIssue] = []
    for f in files:
        text = normalize(f"{f.stem} {f.parent.name}")
        sid_pos, name_pos = _name_positions(idx, text)
        has_sid, has_name = sid_pos >= 0, name_pos >= 0

        bad = (need_sid and not has_sid) or (need_name and not has_name)
        order_bad = False
        if not bad and order and has_sid and has_name:
            if order == "sid_first" and sid_pos > name_pos:
                order_bad = True
            elif order == "name_first" and name_pos > sid_pos:
                order_bad = True
            bad = order_bad

        if bad:
            hits = _find_people(idx, text, True, "文件名")
            person = hits[0][0] if hits else None
            issues.append(NameIssue(f, has_sid, has_name, person, order_bad))
    return issues
