#!/usr/bin/env python3
from __future__ import annotations

import os
import base64
import difflib
import html
import io
import json
import re
import site
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
from urllib.parse import quote, urlencode
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET

# 防止误用用户级站点包或错误工作目录导致 numpy 加载异常
os.environ["PYTHONNOUSERSITE"] = "1"
USER_SITE = site.getusersitepackages()
sys.path = [p for p in sys.path if USER_SITE not in p]

BASE_DIR = Path(__file__).resolve().parent
try:
    os.chdir(BASE_DIR)
except Exception:
    pass


def _write_fatal_log(exc: BaseException) -> Path | None:
    try:
        base = Path.home() / "Library" / "Logs" / "WECARE 产品解决方案"
        base.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%Y%m%d-%H%M%S")
        p = base / f"fatal-{ts}.log"
        p.write_text(
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            encoding="utf-8",
        )
        return p
    except Exception:
        return None


def _show_fatal_dialog(title: str, message: str) -> None:
    if sys.platform != "darwin":
        return
    try:
        import subprocess

        title_lit = json.dumps(title, ensure_ascii=False)
        message_lit = json.dumps(message, ensure_ascii=False)
        subprocess.run(
            [
                "osascript",
                "-e",
                f"display dialog {message_lit} with title {title_lit} buttons {{\"OK\"}} default button \"OK\"",
            ],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass


def _purge_numpy_shadows() -> None:
    """Remove sys.path entries that shadow numpy with an incomplete copy.

    常见触发场景：从 PyInstaller 的 dist/_internal 目录（含不完整的 numpy）
    启动 Python/Streamlit，导致优先导入到缺少 numpy/__config__.py 的目录。
    """

    try:
        site_packages = [Path(p).resolve() for p in site.getsitepackages()]
    except Exception:
        site_packages = []

    def is_under_site_packages(path: Path) -> bool:
        for sp in site_packages:
            try:
                path.relative_to(sp)
                return True
            except Exception:
                continue
        return False

    cleaned: List[str] = []
    for entry in list(sys.path):
        try:
            path = Path(entry).resolve() if entry else Path.cwd().resolve()
        except Exception:
            cleaned.append(entry)
            continue

        if is_under_site_packages(path):
            cleaned.append(entry)
            continue

        numpy_dir = path / "numpy"
        if (
            (numpy_dir / "__init__.py").exists()
            and not (numpy_dir / "__config__.py").exists()
        ):
            continue

        cleaned.append(entry)

    sys.path = cleaned


_purge_numpy_shadows()


def resource_path(relative: str) -> Path:
    """Return resource path for packaged or dev mode."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).parent / relative


def _bundle_parent_dir() -> Path | None:
    if not getattr(sys, "frozen", False):
        return None

    exe = Path(sys.executable).resolve()
    for parent in exe.parents:
        if parent.suffix == ".app":
            return parent.parent

    return exe.parent


def _find_latest_excel(search_dir: Path, patterns: List[str]) -> Path | None:
    candidates: List[Path] = []
    try:
        for pattern in patterns:
            candidates.extend(
                [
                    p
                    for p in search_dir.glob(pattern)
                    if p.is_file() and not p.name.startswith("~$")
                ]
            )
    except Exception:
        return None

    if not candidates:
        return None

    try:
        return max(candidates, key=lambda p: p.stat().st_mtime)
    except Exception:
        return sorted(candidates)[-1]


def resolve_excel_path() -> Path | None:
    env_url = os.getenv("DESIGN_EXCEL_URL")
    if env_url:
        downloaded = _maybe_download_excel(env_url)
        if downloaded and downloaded.exists():
            return downloaded

    env_path = os.getenv("DESIGN_EXCEL")
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p

    patterns = ["产品配方设计*.xlsx"]

    search_dirs: List[Path] = []
    bundle_parent = _bundle_parent_dir()
    if bundle_parent:
        search_dirs.append(bundle_parent)
        search_dirs.append(Path(sys.executable).resolve().parent)

    search_dirs.append(BASE_DIR)

    for search_dir in search_dirs:
        latest = _find_latest_excel(search_dir, patterns)
        if latest:
            return latest

    p = resource_path("产品配方设计最新.xlsx")
    if p.exists():
        return p

    return None


def _excel_cache_dir() -> Path:
    for base in (Path.home() / ".cache", Path.home() / "Library" / "Caches"):
        try:
            base.mkdir(parents=True, exist_ok=True)
            d = base / "wecare-solution"
            d.mkdir(parents=True, exist_ok=True)
            return d
        except Exception:
            continue
    d = Path(tempfile.gettempdir()) / "wecare-solution"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _excel_cache_key(url: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", url)[:120] or "excel"


def _maybe_download_excel(url: str) -> Path | None:
    cache_dir = _excel_cache_dir()
    key = _excel_cache_key(url)
    xlsx_path = cache_dir / f"{key}.xlsx"
    meta_path = cache_dir / f"{key}.json"

    headers: Dict[str, str] = {}
    try:
        if meta_path.exists():
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            if meta.get("etag"):
                headers["If-None-Match"] = meta["etag"]
            if meta.get("last_modified"):
                headers["If-Modified-Since"] = meta["last_modified"]
    except Exception:
        headers = {}

    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            if resp.status == 304 and xlsx_path.exists():
                return xlsx_path
            if resp.status >= 400:
                return xlsx_path if xlsx_path.exists() else None

            content = resp.read()
            xlsx_path.write_bytes(content)

            try:
                new_meta = {
                    "url": url,
                    "etag": resp.headers.get("ETag", ""),
                    "last_modified": resp.headers.get("Last-Modified", ""),
                    "fetched_at": time.time(),
                }
                meta_path.write_text(json.dumps(new_meta, ensure_ascii=False, indent=2), encoding="utf-8")
            except Exception:
                pass
            return xlsx_path
    except Exception:
        return xlsx_path if xlsx_path.exists() else None


# 用于开发/打包时的本地默认路径；在线托管时会在 main() 内按需刷新
EXCEL_PATH = resolve_excel_path()
DOCS_DIR = resource_path("功能说明")
LOGO_PATH = resource_path("wecare_logo.png")
LOGO_ICON_PATH = resource_path("wecare_logo_icon_1024.png")
LOGO_SVG_PATH = resource_path("Final/logo.svg")
HERO_ART_PATH = resource_path("Final/hero.jpg")
PPT_SOLUTIONS_PATH = resource_path("Final/43 Solutions解决方案中文版20260130.pptx")
PPT_SOLUTIONS_EN_PATH = resource_path("Final/43 Solutions解决方案英文版20260203.pptx")
PPT_FORMULA_PATH = resource_path("Final/Formula&Solution.pptx")
PDF_SOLUTIONS_PATH = resource_path("Final/43 Solutions解决方案中文版20260130.pdf")
PDF_SOLUTIONS_EN_PATH = resource_path("Final/43 Solutions解决方案英文版20260203.pdf")
CAPSULE_DETAILS_PATH = resource_path("Final/Capsule配方详情.xlsx")
PPT_WECLAC_PATH = resource_path("Final/WecLac.pptx")
PPT_WECPRO_FORMULA_PATH = resource_path("Final/Formula.pptx")
WECLAC_IMAGES_DIR = resource_path("Final/images")
WECLAC_CORE_CODES = {"BLa80", "LRa05", "BL21", "BC99", "Akk11"}
CLINICAL_DATA_PATH = resource_path("Final/Clinicaldata0201.xlsx")
WECLAC_SCI_NAMES: Dict[str, str] = {
    "BLa80": "Bifidobacterium animalis subsp. lactis",
    "BLa36": "Bifidobacterium animalis subsp. lactis",
    "LRa05": "Lacticaseibacillus rhamnosus",
    "BL21": "Bifidobacterium longum subsp. longum",
    "BC99": "Weizmannia coagulans",
    "Akk11": "Akkermansia muciniphila",
    "LA85": "Lactobacillus acidophilus",
    "LC86": "Lacticaseibacillus paracasei",
    "BBr60": "Bifidobacterium breve",
    "PA53": "Pediococcus acidilactici",
    "Lp05": "Lactiplantibacillus plantarum",
    "Lp18": "Lactiplantibacillus plantarum",
    "BC179": "Weizmannia coagulans",
    "LCr86": "Lactobacillus crispatus",
    "LR08": "Limosilactobacillus reuteri",
}

# More strain codes used across Solutions decks
_STRAIN_SCI_NAMES: Dict[str, str] = {
    **WECLAC_SCI_NAMES,
    "Lp90": "Lactiplantibacillus plantarum",
    "LS97": "Ligilactobacillus salivarius",
    # Used in Formula / Solutions decks
    "BAC30": "Bifidobacterium adolescentis",
    "BI45": "Bifidobacterium longum subsp. infantis",
}

# 延后导入重依赖库，避免环境变量未生效
import pandas as pd  # noqa: E402
import streamlit as st  # noqa: E402


@st.cache_data(ttl=300)
def fetch_remote_excel(url: str) -> str:
    """下载远程 Excel（带 5 分钟 TTL），返回本地缓存路径字符串。"""
    p = _maybe_download_excel(url)
    return str(p) if p else ""


def _normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


_CLINICAL_SEGMENT_SPLIT_RE = re.compile(r"[,，;\n\r；]+")
_CLINICAL_TOKEN_SPLIT_RE = re.compile(r"[\s/／、]+")
_EXCIPIENT_SPLIT_RE = re.compile(r"[,，;\n\r；、]+")


def _split_excipients(excipients: str) -> Tuple[str, str]:
    """把“辅料”拆分为：益生元、其他辅料。"""
    text = (excipients or "").strip()
    if not text:
        return "", ""

    # Excel 已经按两行结构化（优先使用）
    prebiotic_label = "益生元："
    other_label = "其他辅料："
    if prebiotic_label in text or other_label in text:
        prebiotics: List[str] = []
        others: List[str] = []
        for line in [l.strip() for l in text.splitlines() if l.strip()]:
            if line.startswith(prebiotic_label):
                prebiotics.append(line[len(prebiotic_label) :].strip())
            elif line.startswith(other_label):
                others.append(line[len(other_label) :].strip())
            else:
                others.append(line)
        return "、".join([x for x in prebiotics if x]), "、".join([x for x in others if x])

    tokens = [t.strip() for t in _EXCIPIENT_SPLIT_RE.split(text) if t.strip()]
    prebiotic_keywords = (
        "低聚果糖",
        "低聚半乳糖",
        "菊粉",
        "抗性糊精",
        "低聚异麦芽糖",
        "低聚木糖",
        "低聚甘露糖",
        "棉子糖",
        "2'-岩藻糖基乳糖",
        "2'-FL",
        "GOS",
        "FOS",
        "HMO",
    )

    prebiotics: List[str] = []
    others: List[str] = []
    for token in tokens:
        if any(k in token for k in prebiotic_keywords):
            prebiotics.append(token)
        else:
            others.append(token)

    return "、".join(prebiotics), "、".join(others)


def _split_clinical_tokens(text: str) -> List[str]:
    if not text:
        return []
    normalized = (
        text.replace("：", ":")
        .replace("\t", " ")
        .replace("，", ",")
        .replace("；", ";")
        .replace("、", " ")
    )
    return [t for t in _CLINICAL_TOKEN_SPLIT_RE.split(normalized) if t]


def _format_clinical_regs_markdown(text: str) -> str:
    """把临床号按“菌株: 临床号”分组；无菌株开头的行自动并入上一组。"""
    if not text:
        return ""

    groups: List[Dict[str, List[str]]] = []
    segments = [s.strip() for s in _CLINICAL_SEGMENT_SPLIT_RE.split(text) if s.strip()]
    for seg in segments:
        seg = seg.replace("：", ":").strip()
        if ":" in seg:
            label, rest = seg.split(":", 1)
            current_label = label.strip()
            current_ids: List[str] = []
            for token in _split_clinical_tokens(rest.strip()):
                token = token.replace("：", ":").strip()
                if ":" in token:
                    next_label, next_rest = token.split(":", 1)
                    next_label = next_label.strip()
                    if next_label:
                        if current_ids:
                            groups.append({"label": current_label, "ids": current_ids})
                        current_label = next_label
                        current_ids = []
                        next_rest = next_rest.strip()
                        if next_rest:
                            current_ids.extend(_split_clinical_tokens(next_rest))
                        continue
                if token:
                    current_ids.append(token)
            if current_ids:
                groups.append({"label": current_label, "ids": current_ids})
            continue

        tokens = _split_clinical_tokens(seg)
        if not tokens:
            continue
        if groups:
            groups[-1]["ids"].extend(tokens)
        else:
            groups.append({"label": "", "ids": tokens})

    lines: List[str] = []
    for g in groups:
        label = g.get("label", "").strip()
        ids: List[str] = []
        seen: set[str] = set()
        for t in g.get("ids", []):
            if t in seen:
                continue
            seen.add(t)
            ids.append(t)
        if not ids:
            continue
        ids_md = ", ".join(f"`{x}`" for x in ids)
        if label:
            lines.append(f"- `{label}`: {ids_md}")
        else:
            lines.append(f"- {ids_md}")

    return "\n".join(lines)


def _parse_clinical_regs_entries(text: str) -> List[Tuple[str, List[str]]]:
    """Parse scenario-level clinical regs text into [(label, [ids...])]."""
    if not text:
        return []

    groups: List[Dict[str, List[str]]] = []
    segments = [s.strip() for s in _CLINICAL_SEGMENT_SPLIT_RE.split(text) if s.strip()]
    for seg in segments:
        seg = seg.replace("：", ":").strip()
        if ":" in seg:
            label, rest = seg.split(":", 1)
            current_label = label.strip()
            current_ids: List[str] = []
            for token in _split_clinical_tokens(rest.strip()):
                token = token.replace("：", ":").strip()
                if ":" in token:
                    next_label, next_rest = token.split(":", 1)
                    next_label = next_label.strip()
                    if next_label:
                        if current_ids:
                            groups.append({"label": current_label, "ids": current_ids})
                        current_label = next_label
                        current_ids = []
                        next_rest = next_rest.strip()
                        if next_rest:
                            ids = re.findall(r"(NCT\d+|ChiCTR\d+)", next_rest, flags=re.IGNORECASE)
                            current_ids.extend(ids if ids else [next_rest])
                        continue
                if token:
                    ids = re.findall(r"(NCT\d+|ChiCTR\d+)", token, flags=re.IGNORECASE)
                    if ids:
                        current_ids.extend(ids)
                    else:
                        current_ids.append(token)
            if current_ids:
                groups.append({"label": current_label, "ids": current_ids})
            continue

        tokens = _split_clinical_tokens(seg)
        if not tokens:
            continue
        parsed_ids: List[str] = []
        for token in tokens:
            ids = re.findall(r"(NCT\d+|ChiCTR\d+)", token, flags=re.IGNORECASE)
            if ids:
                parsed_ids.extend(ids)
            else:
                parsed_ids.append(token)
        if groups:
            groups[-1]["ids"].extend(parsed_ids)
        else:
            groups.append({"label": "", "ids": parsed_ids})

    out: List[Tuple[str, List[str]]] = []
    for g in groups:
        label = (g.get("label") or "").strip()
        uniq_ids: List[str] = []
        seen: set[str] = set()
        for one in g.get("ids", []):
            rid = str(one or "").strip()
            if not rid:
                continue
            rid = re.sub(r"[,;；，]+$", "", rid)
            if rid in seen:
                continue
            seen.add(rid)
            uniq_ids.append(rid)
        if uniq_ids:
            out.append((label, uniq_ids))
    return out


@st.cache_data
def load_solution_design(
    excel_path: str,
    _cache_buster: float | None = None,
) -> Tuple[
    Dict[str, Dict[str, List[Dict[str, str]]]],
    Dict[str, Dict[str, Dict[str, str]]],
    List[str],
    Dict[str, List[str]],
]:
    """读取产品配方设计表（Sheet1）。

    返回：
    - mapping: {功能方向: {细分方向: [{菌株, 临床证据}]}}
    - meta: {功能方向: {细分方向: {solution, excipients, clinical_regs}}}
    - main_order: 功能方向顺序
    - sub_order: {功能方向: [细分方向顺序]}
    """

    raw = pd.read_excel(excel_path, sheet_name="Sheet1", header=None)
    if raw.shape[0] < 3 or raw.shape[1] < 2:
        return {}, {}, [], {}

    header_main = raw.iloc[0].ffill()
    header_sub = raw.iloc[1].ffill()

    special_labels = ["菌株应用解决方案", "益生元", "其他辅料", "辅料", "相关临床注册号"]
    special_rows: Dict[str, int] = {}
    for idx in range(raw.shape[0]):
        label = raw.iloc[idx, 0]
        if isinstance(label, str) and label.strip() in special_labels:
            special_rows[label.strip()] = idx

    end_data = min(special_rows.values()) if special_rows else raw.shape[0]
    data = raw.iloc[2:end_data]

    mapping: Dict[str, Dict[str, List[Dict[str, str]]]] = {}
    meta: Dict[str, Dict[str, Dict[str, str]]] = {}
    main_order: List[str] = []
    sub_order: Dict[str, List[str]] = {}

    for col_idx in range(1, raw.shape[1]):
        main = _normalize_text(header_main[col_idx])
        sub = _normalize_text(header_sub[col_idx])
        if not main or not sub:
            continue

        if main not in main_order:
            main_order.append(main)
        sub_order.setdefault(main, []).append(sub)

        mapping.setdefault(main, {}).setdefault(sub, [])
        meta.setdefault(main, {}).setdefault(sub, {})

        sol = _normalize_text(
            raw.iloc[special_rows.get("菌株应用解决方案", -1), col_idx]
            if "菌株应用解决方案" in special_rows
            else ""
        )
        pre = _normalize_text(
            raw.iloc[special_rows.get("益生元", -1), col_idx] if "益生元" in special_rows else ""
        )
        other_exc = _normalize_text(
            raw.iloc[special_rows.get("其他辅料", -1), col_idx]
            if "其他辅料" in special_rows
            else ""
        )
        exc = _normalize_text(
            raw.iloc[special_rows.get("辅料", -1), col_idx] if "辅料" in special_rows else ""
        )
        regs = _normalize_text(
            raw.iloc[special_rows.get("相关临床注册号", -1), col_idx]
            if "相关临床注册号" in special_rows
            else ""
        )
        meta[main][sub] = {
            "solution": sol,
            "prebiotics": pre,
            "other_excipients": other_exc,
            "excipients": exc,
            "clinical_regs": regs,
        }

        for _, row in data.iterrows():
            strain = _normalize_text(row.iloc[0])
            value = _normalize_text(row.iloc[col_idx])
            if not strain or not value:
                continue
            mapping[main][sub].append({"菌株": strain, "临床证据": value})

    return mapping, meta, main_order, sub_order


@st.cache_data
def load_product_overview(
    excel_path: str, _cache_buster: float | None = None
) -> Dict[str, Dict[str, str]]:
    """读取产品配方设计表（Sheet2）。

    返回：{功能方向: {name, core_formula, clinical_regs}}
    """

    try:
        df = pd.read_excel(excel_path, sheet_name="Sheet2")
    except ValueError:
        try:
            df = pd.read_excel(excel_path, sheet_name=1)
        except Exception:
            return {}
    except Exception:
        return {}

    func_col = None
    for candidate in ("功能", "功能方向"):
        if candidate in df.columns:
            func_col = candidate
            break

    name_col = "名称" if "名称" in df.columns else None
    formula_col = "核心配方" if "核心配方" in df.columns else None
    regs_col = "临床注册号" if "临床注册号" in df.columns else None

    if not func_col:
        return {}

    overview: Dict[str, Dict[str, str]] = {}
    for _, row in df.iterrows():
        func = _normalize_text(row.get(func_col))
        if not func:
            continue

        overview[func] = {
            "name": _normalize_text(row.get(name_col)) if name_col else "",
            "core_formula": _normalize_text(row.get(formula_col)) if formula_col else "",
            "clinical_regs": _normalize_text(row.get(regs_col)) if regs_col else "",
        }

    return overview


_PPT_DRAWING_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _pptx_extract_paragraph_lines(xml_bytes: bytes) -> List[str]:
    """从 slide.xml 提取段落文本（按 PPT 中的段落聚合）。"""
    root = ET.fromstring(xml_bytes)
    lines: List[str] = []
    for p in root.findall(".//a:p", _PPT_DRAWING_NS):
        text = "".join((t.text or "") for t in p.findall(".//a:t", _PPT_DRAWING_NS))
        text = text.replace("\u00a0", " ").strip()
        if text:
            lines.append(text)
    return lines


def _pptx_slide_paths(z: zipfile.ZipFile) -> Dict[int, str]:
    slide_paths = [
        n
        for n in z.namelist()
        if n.startswith("ppt/slides/slide") and n.endswith(".xml")
    ]

    out: Dict[int, str] = {}
    for n in slide_paths:
        m = re.search(r"ppt/slides/slide(\d+)\.xml$", n)
        if not m:
            continue
        out[int(m.group(1))] = n
    return out


def _ppt_solution_title_from_lines(lines: List[str]) -> str:
    if not lines:
        return ""

    # Pattern 1: "Solution | Title" / "Solution ｜ Title"
    for raw in lines:
        line = str(raw or "").strip()
        if not line:
            continue
        if re.search(r"\bSolution\b", line, flags=re.IGNORECASE) and ("|" in line or "｜" in line):
            parts = re.split(r"[|｜]", line)
            if len(parts) >= 2:
                cand = parts[-1].strip()
                if cand:
                    return cand

    # Pattern 2: standalone "Solution" then next line is title
    for i, raw in enumerate(lines[:-1]):
        line = str(raw or "").strip()
        if re.fullmatch(r"solution", line, flags=re.IGNORECASE):
            cand = str(lines[i + 1] or "").strip()
            if cand and not re.fullmatch(r"solutions?", cand, flags=re.IGNORECASE):
                return cand

    return ""


def _pdf_solution_title_from_text(text: str) -> str:
    """Extract a Solution title from PDF page text (EN fallback when PPT text is not extractable)."""
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    if not lines:
        return ""

    for line in lines:
        if re.search(r"\bSolution\b", line, flags=re.IGNORECASE) and ("|" in line or "｜" in line):
            parts = re.split(r"[|｜]", line)
            if len(parts) >= 2:
                cand = parts[-1].strip()
                if cand:
                    return cand

    for i, line in enumerate(lines[:-1]):
        if line.strip().lower() == "solution":
            return lines[i + 1].strip()

    return ""


@st.cache_data
def load_pdf_solution_titles(
    pdf_path: str, _cache_buster: float | None = None
) -> Dict[int, str]:
    """Return {page_no(1-based): title} extracted from the Solutions PDF."""
    p = Path(pdf_path)
    if not p.exists():
        return {}
    try:
        import fitz  # type: ignore[import-not-found]
    except Exception:
        return {}

    titles: Dict[int, str] = {}
    doc = fitz.open(str(p))
    try:
        total = int(getattr(doc, "page_count", len(doc)))
        for i in range(total):
            page = doc.load_page(i)
            txt = page.get_text("text") or ""
            title = _pdf_solution_title_from_text(txt)
            if title:
                titles[i + 1] = title
    finally:
        doc.close()
    return titles


@st.cache_data
def load_pdf_solution_start_pages(
    pdf_path: str, _cache_buster: float | None = None
) -> List[Tuple[int, str]]:
    """Extract ordered solution start pages as [(page_no, title), ...] from PDF."""
    titles = load_pdf_solution_titles(pdf_path, _cache_buster)
    if not titles:
        return []

    starts: List[Tuple[int, str]] = []
    prev_norm = ""
    for page_no in sorted(titles.keys()):
        title = str(titles.get(page_no, "") or "").strip()
        if not title:
            continue
        norm = _normalize_match_key(title).lower()
        if norm and norm == prev_norm:
            continue
        starts.append((page_no, title))
        prev_norm = norm
    return starts


def _normalize_match_key(text: str) -> str:
    return (
        (text or "")
        .replace("：", ":")
        .replace("（", "(")
        .replace("）", ")")
        .replace("，", ",")
        .replace("；", ";")
        .replace("、", " ")
        .strip()
        .replace(" ", "")
    )


def _clean_ui_key(v: object) -> str:
    """Normalize UI keys (category/scenario labels) to avoid hidden whitespace issues."""
    s = str(v or "")
    s = s.replace("\u00a0", " ").replace("\u200b", "")
    return s.strip()


@st.cache_data
def load_ppt_solution_deck(
    pptx_path: str, _cache_buster: float | None = None
) -> Dict[str, Dict[str, object]]:
    """解析“43 Solutions...”PPTX，返回 {title: {overview_lines, evidence_lines, slide_no}}。"""
    p = Path(pptx_path)
    if not p.exists():
        return {}

    with zipfile.ZipFile(p) as z:
        slide_map = _pptx_slide_paths(z)
        if not slide_map:
            return {}

        max_slide_no = max(slide_map.keys())
        solutions: Dict[str, Dict[str, object]] = {}

        slide_no = 1
        while slide_no <= max_slide_no:
            if slide_no % 2 == 0:
                slide_no += 1
                continue

            overview_path = slide_map.get(slide_no)
            evidence_path = slide_map.get(slide_no + 1)
            if not overview_path or not evidence_path:
                slide_no += 2
                continue

            overview_lines = _pptx_extract_paragraph_lines(z.read(overview_path))
            evidence_lines = _pptx_extract_paragraph_lines(z.read(evidence_path))
            title = _ppt_solution_title_from_lines(overview_lines) or _ppt_solution_title_from_lines(
                evidence_lines
            )
            if title:
                solutions[title] = {
                    "overview_lines": overview_lines,
                    "evidence_lines": evidence_lines,
                    "slide_no": slide_no,
                }

            slide_no += 2

        return solutions


@st.cache_data
def load_ppt_solution_start_slides(
    pptx_path: str, _cache_buster: float | None = None
) -> List[Tuple[int, str]]:
    """Extract ordered solution start slides as [(slide_no, title), ...].

    This scanner reads all slides and keeps the first title of each consecutive
    duplicate pair, so it works even when EN decks include intro/summary pages
    and page numbering does not align with CN decks.
    """
    p = Path(pptx_path)
    if not p.exists():
        return []

    all_hits: List[Tuple[int, str]] = []
    with zipfile.ZipFile(p) as z:
        slide_map = _pptx_slide_paths(z)
        for slide_no in sorted(slide_map.keys()):
            try:
                lines = _pptx_extract_paragraph_lines(z.read(slide_map[slide_no]))
            except Exception:
                continue
            title = _ppt_solution_title_from_lines(lines)
            if title:
                all_hits.append((slide_no, title.strip()))

    if not all_hits:
        return []

    starts: List[Tuple[int, str]] = []
    prev_norm = ""
    for slide_no, title in all_hits:
        norm = _normalize_match_key(title).lower()
        if norm and norm == prev_norm:
            continue
        starts.append((slide_no, title))
        prev_norm = norm

    return starts


_FORMULA_SLIDE_TO_DIRECTION = {
    2: "女性健康",
    3: "情绪健康",
    4: "代谢健康",
    5: "胃肠健康",
    6: "免疫健康",
    7: "婴童健康",
    8: "口腔健康",
}


def _parse_formula_slide_scenarios(lines: List[str]) -> List[str]:
    header = {"应用场景", "临床菌配方", "菌株配方", "临床验证及注册号"}
    scenarios: List[str] = []
    for line in lines:
        if line in header:
            continue
        if line.startswith("人类营养与健康"):
            break
        if "+" in line or "NCT" in line or "ChiCTR" in line or ":" in line:
            continue
        if not re.search(r"[\u4e00-\u9fff]", line):
            continue
        scenarios.append(line.strip())
    return scenarios


@st.cache_data
def load_formula_scenarios(
    pptx_path: str, _cache_buster: float | None = None
) -> Dict[str, List[str]]:
    """解析“Formula&Solution.pptx”中各功能方向的“应用场景”列表。"""
    p = Path(pptx_path)
    if not p.exists():
        return {}

    with zipfile.ZipFile(p) as z:
        slide_map = _pptx_slide_paths(z)
        out: Dict[str, List[str]] = {}
        for slide_no, direction in _FORMULA_SLIDE_TO_DIRECTION.items():
            slide_path = slide_map.get(slide_no)
            if not slide_path:
                continue
            lines = _pptx_extract_paragraph_lines(z.read(slide_path))
            out[direction] = _parse_formula_slide_scenarios(lines)
        return out


@st.cache_data
def load_pptx_slide_lines(
    pptx_path: str, slide_no: int = 1, _cache_buster: float | None = None
) -> List[str]:
    p = Path(pptx_path)
    if not p.exists():
        return []

    with zipfile.ZipFile(p) as z:
        slide_map = _pptx_slide_paths(z)
        slide_path = slide_map.get(int(slide_no))
        if not slide_path:
            return []
        return _pptx_extract_paragraph_lines(z.read(slide_path))


@st.cache_data
def load_wecpro_formula_catalog(
    pptx_path: str, _cache_buster: float | None = None
) -> List[Dict[str, object]]:
    """解析 `Final/Formula.pptx`，返回每个功能方向的 WecPro® Formula 信息。"""
    lines = load_pptx_slide_lines(pptx_path, 1, _cache_buster)
    if not lines:
        return []

    header = {"WecPro Formula", "功能", "商品名", "健康功效", "核心配方"}
    directions = set(_FORMULA_SLIDE_TO_DIRECTION.values())
    clean = [l.strip() for l in lines if l.strip() and l.strip() not in header]

    out: List[Dict[str, object]] = []
    i = 0
    while i < len(clean):
        if clean[i] not in directions:
            i += 1
            continue

        direction = clean[i]
        product = clean[i + 1].strip() if i + 1 < len(clean) else ""
        benefit = clean[i + 2].strip() if i + 2 < len(clean) else ""
        i += 3

        strains: List[str] = []
        while i < len(clean) and clean[i] not in directions:
            strains.append(clean[i])
            i += 1

        out.append(
            {
                "direction": direction,
                "product": product,
                "benefit": benefit,
                "strains": strains,
            }
        )

    return out


@st.cache_data
def load_weclac_catalog(
    pptx_path: str, lang: str = "CN", _cache_buster: float | None = None
) -> Dict[str, object]:
    """解析 `Final/WecLac.pptx` 的菌株表，返回结构化数据。"""
    lang_norm = (lang or "CN").strip().upper()
    slide_no = 2 if lang_norm == "EN" else 1
    lines = load_pptx_slide_lines(pptx_path, slide_no, _cache_buster)
    if not lines:
        return {}

    headers = {
        "产品分类",
        "产品名称",
        "产品特点",
        "临床数量",
        "专利数量",
        "规格",
        "功能方向",
        "WecLac 菌株介绍",
        "Product Category",
        "Strain Name",
        "Strain Highlights",
        "Clinical Studies",
        "Patents",
        "Specification (CFU)",
        "Supported Application Areas",
        "WecLac Strains Introduction",
    }
    clean = [l.strip() for l in lines if l.strip() and l.strip() not in headers]

    product = "WecLac"
    product_type = ""
    core_flag = False
    strains: List[Dict[str, object]] = []

    i = 0
    while i < len(clean):
        token = clean[i]

        if token == "WecLac":
            product = token
            i += 1
            continue
        if token in {"益生菌", "Probiotics"}:
            product_type = token
            i += 1
            continue
        if "核心" in token or "core" in token.lower():
            core_flag = True
            i += 1
            continue

        if i + 4 >= len(clean):
            break

        name = clean[i]
        feature = clean[i + 1]
        clinical = clean[i + 2]
        patent = clean[i + 3]
        spec = clean[i + 4]
        i += 5

        functions: List[Dict[str, str]] = []
        while i < len(clean) and re.match(r"^\d+\.\s*", clean[i]):
            raw_dir = re.sub(r"^\d+\.\s*", "", clean[i]).strip()
            direction = raw_dir
            desc = ""

            if "：" in raw_dir:
                left, right = raw_dir.split("：", 1)
                direction = left.strip() or raw_dir
                desc = right.strip()
                i += 1
            elif ":" in raw_dir:
                left, right = raw_dir.split(":", 1)
                direction = left.strip() or raw_dir
                desc = right.strip()
                i += 1
            elif i + 1 < len(clean) and not re.match(r"^\d+\.\s*", clean[i + 1]):
                desc = clean[i + 1].strip()
                i += 2
            else:
                i += 1
            if direction:
                functions.append({"direction": direction, "desc": desc})

        strains.append(
            {
                "name": name,
                "feature": feature,
                "clinical": clinical,
                "patent": patent,
                "spec": spec,
                "core": bool(core_flag),
                "functions": functions,
            }
        )
        core_flag = False

    # Explicit field fixes for known source typos in WecLac deck.
    # Key: (LANG, code) -> corrected spec text
    spec_overrides: Dict[Tuple[str, str], str] = {
        ("EN", "BC179"): "300B",
    }
    for row in strains:
        name = str(row.get("name", "")).strip()
        _, code = _extract_strain_code(name)
        override_spec = spec_overrides.get((lang_norm, code))
        if override_spec:
            row["spec"] = override_spec

    return {"product": product, "product_type": product_type, "strains": strains}


def _best_title_match(query: str, candidates: List[str]) -> str:
    if not candidates:
        return ""

    q = _normalize_match_key(query).replace("调控", "调节")
    best_title = ""
    best_score = -1.0
    for title in candidates:
        t = _normalize_match_key(title).replace("调控", "调节")
        if not q or not t:
            continue

        if q == t:
            score = 1.0
        elif q in t or t in q:
            score = 0.95
        else:
            score = difflib.SequenceMatcher(None, q, t).ratio()

        if score > best_score:
            best_score = score
            best_title = title

    return best_title


@st.cache_data
def build_scenario_to_solution_title(
    formula_pptx_path: str,
    solutions_pptx_path: str,
    _cache_buster: float | None = None,
) -> Dict[str, str]:
    """把“应用场景”（Formula PPT）映射到“43 Solutions”中的 solution title。"""
    direction_to_scenarios = load_formula_scenarios(formula_pptx_path, _cache_buster)
    solutions = load_ppt_solution_deck(solutions_pptx_path, _cache_buster)
    if not direction_to_scenarios or not solutions:
        return {k: k for k in solutions.keys()}

    # 以 PPT 中的顺序为准：按 slide_no 排序
    solution_titles_ordered = [
        title
        for title, _ in sorted(
            solutions.items(), key=lambda kv: int(kv[1].get("slide_no", 10**9))
        )
    ]

    # 全局一一匹配：避免“某个大类条数变化”导致后续全部串位。
    alias: Dict[str, str] = {}
    remaining = list(solution_titles_ordered)
    for slide_no in sorted(_FORMULA_SLIDE_TO_DIRECTION.keys()):
        direction = _FORMULA_SLIDE_TO_DIRECTION[slide_no]
        scenarios = direction_to_scenarios.get(direction, [])
        if not scenarios:
            continue
        for scen in scenarios:
            picked = _best_title_match(scen, remaining)
            if not picked:
                picked = _best_title_match(scen, solution_titles_ordered)
            if picked:
                alias[scen] = picked
                try:
                    remaining.remove(picked)
                except ValueError:
                    pass

    # Identity mapping for direct hits
    for title in solution_titles_ordered:
        alias.setdefault(title, title)

    # Normalized-key mapping for robustness
    normalized: Dict[str, str] = {}
    for k, v in alias.items():
        normalized[_normalize_match_key(k)] = v
    alias.update(normalized)

    return alias


_PPT_STRAIN_HINTS = (
    "乳杆菌",
    "双歧杆菌",
    "芽孢杆菌",
    "魏茨曼",
    "阿克曼",
    "片球菌",
    "链球菌",
    "Lactobacillus",
    "Bifidobacterium",
    "Akkermansia",
    "Weizmannia",
    "Pediococcus",
    "Lacticaseibacillus",
    "Limosilactobacillus",
    "Lactiplantibacillus",
)

_PPT_EXCIPIENT_HINTS = (
    "核心辅料",
    "其他辅料",
    "辅料",
    "excipients",
    "xcipients",
    "inulin",
    "acacia gum",
    "gum arabic",
    "resistant dextrin",
    "fructo-oligosaccharides",
    "potato starch",
    "starch",
    "dextrin",
    "酵母",
    "菊粉",
    "阿拉伯胶",
    "抗性糊精",
    "低聚果糖",
    "淀粉",
)


def _contains_ppt_strain_code(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    if re.search(r"\b(NCT\d+|ChiCTR\d+)\b", s, flags=re.I):
        return False
    return bool(re.search(r"(?<![A-Za-z])[A-Za-z]{1,6}\d{2,3}(?![A-Za-z])", s))


def _is_ppt_strain_line(line: str) -> bool:
    s = (line or "").strip()
    if not _contains_ppt_strain_code(s):
        return False
    if any(hint in s for hint in _PPT_STRAIN_HINTS):
        return True
    # Compact blend formats: LRa05+LCr86+LR08 / LRa05,LCr86,LR08
    return bool(
        re.search(
            r"(?:[A-Za-z]{1,6}\d{2,3}\s*[+＋/／,，;；]){1,}[A-Za-z]{1,6}\d{2,3}",
            s,
        )
    )


def _is_ppt_excipients_labeled_line(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    if re.match(r"^(核心辅料|其他辅料|辅料)\s*[:：]", s):
        return True
    if re.match(
        r"^(Key\s+Excipients?|Other\s+(?:Excipients?|Xcipients?)|Excipients?)\s*[:：]",
        s,
        flags=re.I,
    ):
        return True
    return False


def _is_ppt_excipients_continuation_line(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    if _is_ppt_trial_line(s):
        return False
    if _is_ppt_strain_line(s):
        return False
    if re.search(r"\b\d+\s*菌株\b", s) or re.search(r"\b\d+\s*Strains?\b", s, flags=re.I):
        return False
    low = s.lower()
    if any(h in low for h in _PPT_EXCIPIENT_HINTS):
        return True
    return False


def _is_ppt_trial_line(line: str) -> bool:
    if re.search(r"\b(NCT\d+|ChiCTR\d+)\b", line):
        return True
    if re.match(r"^[A-Za-z0-9+/_-]+:\s*", line):
        return True
    return False


def _is_ppt_noise_line(line: str) -> bool:
    if line == "Partial data shown. More data available...":
        return True
    if re.fullmatch(r"[A-Za-z0-9_.-]+", line):
        return True
    if re.fullmatch(r"[A-Za-z0-9 .-]+", line) and len(line) <= 18:
        return True
    return False


def _parse_ppt_overview(lines: List[str]) -> Dict[str, object]:
    title = _ppt_solution_title_from_lines(lines)
    try:
        sol_idx = lines.index("Solution")
    except ValueError:
        sol_idx = -1

    content_lines = []
    if sol_idx > 0:
        content_lines.extend(lines[:sol_idx])
    if sol_idx >= 0:
        after = lines[sol_idx + 1 :]
        if after and title and after[0] == title:
            after = after[1:]
        content_lines.extend(after)
    else:
        content_lines = list(lines)

    meta_lines = {
        "临床研究",
        "科学支持",
        "研究成果",
        "核心功能",
        "CLINICAL STUDIES",
        "Clinical Studies",
        "Scientific Support",
        "RESEARCH OUTCOME",
        "Research Outcome",
        "Functionality",
        "Functions",
        "Function",
    }
    specs: List[str] = []
    strains: List[str] = []
    excipients: List[str] = []
    trials: List[str] = []
    highlights: List[str] = []
    in_excipients_block = False

    for line in content_lines:
        if not line or line in meta_lines or line == "Solution" or line == title:
            continue

        if _is_ppt_excipients_labeled_line(line):
            excipients.append(line)
            in_excipients_block = True
            continue
        if in_excipients_block and _is_ppt_excipients_continuation_line(line):
            excipients.append(line)
            continue
        in_excipients_block = False

        if (
            re.search(r"\b\d+\s*菌株\b", line)
            or re.search(r"\b\d+\s*Strains?\b", line, flags=re.I)
            or line in {"粉剂/胶囊", "辅料可选", "Powder / Capsule Form", "Powder/Capsule Form"}
            or ("optional" in line.lower() and "excipient" in line.lower())
        ):
            specs.append(line)
            continue

        if _is_ppt_trial_line(line):
            trials.append(line)
            continue

        if _is_ppt_strain_line(line):
            strains.append(line)
            continue

        highlights.append(line)

    return {
        "title": title,
        "specs": specs,
        "strains": strains,
        "excipients": excipients,
        "trials": trials,
        "highlights": highlights,
    }


def _parse_ppt_evidence(lines: List[str]) -> Dict[str, object]:
    title = _ppt_solution_title_from_lines(lines)
    try:
        sol_idx = lines.index("Solution")
    except ValueError:
        sol_idx = -1

    content = lines[sol_idx + 1 :] if sol_idx >= 0 else list(lines)
    if content and title and content[0] == title:
        content = content[1:]

    bullets: List[str] = []
    dois: List[str] = []
    for line in content:
        if not line or line in {"研究成果", "临床研究", "科学支持", "核心功能"}:
            continue
        if _is_ppt_noise_line(line):
            continue
        if line.startswith("DOI:"):
            dois.append(line.replace("DOI:", "").strip())
            continue
        bullets.append(line)

    return {"title": title, "bullets": bullets, "dois": dois}


def resolve_solutions_pptx_path(lang: str = "CN") -> Path | None:
    """返回 Solutions PPTX 路径（支持 CN/EN）。"""
    lang_norm = (lang or "CN").strip().upper()
    env_key = "DESIGN_SOLUTIONS_PPTX_EN" if lang_norm == "EN" else "DESIGN_SOLUTIONS_PPTX"
    env_path = os.getenv(env_key, "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p

    if lang_norm == "EN" and PPT_SOLUTIONS_EN_PATH.exists():
        return PPT_SOLUTIONS_EN_PATH
    if lang_norm != "EN" and PPT_SOLUTIONS_PATH.exists():
        return PPT_SOLUTIONS_PATH

    search_dir = BASE_DIR / "Final"
    if not search_dir.exists():
        return None

    candidates = [
        p
        for p in search_dir.glob("*.pptx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    if not candidates:
        return None

    if lang_norm == "EN":
        for p in candidates:
            if "英文" in p.name or "English" in p.name:
                return p
    else:
        for p in candidates:
            if "中文" in p.name or "中文版" in p.name:
                return p

    for p in candidates:
        if "Solutions" in p.name or "解决方案" in p.name:
            return p

    return sorted(candidates)[0]


def resolve_solutions_pdf_path(lang: str = "CN") -> Path | None:
    """返回 Solutions PDF 路径（支持 CN/EN）。"""
    lang_norm = (lang or "CN").strip().upper()
    env_key = "DESIGN_SOLUTIONS_PDF_EN" if lang_norm == "EN" else "DESIGN_SOLUTIONS_PDF"
    env_path = os.getenv(env_key, "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p

    if lang_norm == "EN" and PDF_SOLUTIONS_EN_PATH.exists():
        return PDF_SOLUTIONS_EN_PATH
    if lang_norm != "EN" and PDF_SOLUTIONS_PATH.exists():
        return PDF_SOLUTIONS_PATH

    search_dir = BASE_DIR / "Final"
    if search_dir.exists():
        candidates = [p for p in search_dir.glob("*.pdf") if p.is_file()]
        if lang_norm == "EN":
            for p in candidates:
                if "英文" in p.name or "English" in p.name:
                    return p
        else:
            for p in candidates:
                if "中文" in p.name or "中文版" in p.name:
                    return p
        for p in candidates:
            if "Solutions" in p.name or "解决方案" in p.name:
                return p
        if candidates:
            return sorted(candidates)[0]

    return None


def resolve_capsule_details_path() -> Path | None:
    env_path = os.getenv("DESIGN_CAPSULE_XLSX", "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p

    if CAPSULE_DETAILS_PATH.exists():
        return CAPSULE_DETAILS_PATH

    search_dir = BASE_DIR / "Final"
    if search_dir.exists():
        candidates = [
            p
            for p in search_dir.glob("*.xlsx")
            if p.is_file()
            and not p.name.startswith("~$")
            and ("Capsule" in p.name or "胶囊" in p.name)
        ]
        if candidates:
            return sorted(candidates)[0]

    return None


def resolve_clinical_data_path() -> Path | None:
    env_path = os.getenv("DESIGN_CLINICAL_XLSX", "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p

    if CLINICAL_DATA_PATH.exists():
        return CLINICAL_DATA_PATH

    search_dir = BASE_DIR / "Final"
    if search_dir.exists():
        candidates = [
            p
            for p in search_dir.glob("Clinicaldata*.xlsx")
            if p.is_file() and not p.name.startswith("~$")
        ]
        if candidates:
            return sorted(candidates)[0]

    return None


@st.cache_data
def load_clinical_article_links(
    xlsx_path: str, _cache_buster: float | None = None
) -> Dict[str, str]:
    """读取 Clinicaldata*.xlsx，返回 {注册号: SCI 链接}。"""
    p = Path(xlsx_path)
    if not p.exists():
        return {}

    try:
        df = pd.read_excel(p, sheet_name=0)
    except Exception:
        return {}

    # 兼容列名变化
    id_col = "注册号" if "注册号" in df.columns else None
    url_col = "SCI 网页超链接" if "SCI 网页超链接" in df.columns else None
    if not id_col or not url_col:
        return {}

    out: Dict[str, str] = {}
    for _, row in df.iterrows():
        reg_id = _normalize_text(row.get(id_col)).replace(" ", "")
        url = _normalize_text(row.get(url_col)).strip()
        if not reg_id or not url:
            continue
        if not url.startswith(("http://", "https://")):
            continue
        out.setdefault(reg_id, url)
    return out


def _detect_capsule_spec_blocks(raw: pd.DataFrame) -> List[Tuple[int, str]]:
    """在 Capsule 配方详情表中识别“规格块”起始行（例如：0# 胶囊 120B / Capsule 120B）。"""
    blocks: List[Tuple[int, str]] = []
    max_row = min(int(raw.shape[0]), 30)
    for r in range(2, max_row):
        label = _normalize_text(raw.iloc[r, 0])
        if not label:
            continue
        if "胶囊" not in label and "capsule" not in label.lower():
            continue
        if not re.search(r"\d+\s*B", label, flags=re.IGNORECASE):
            continue
        if r + 2 >= raw.shape[0]:
            continue
        blocks.append((r, label))
    return blocks[:3]


@st.cache_data
def load_capsule_details(
    xlsx_path: str, lang: str = "CN", _cache_buster: float | None = None
) -> Dict[str, Dict[str, Dict[str, object]]]:
    """读取 Capsule配方详情.xlsx。

    返回：{功能方向: {产品解决方案: {scenario, direction, specs:[{spec, clinical, excipients, total}]}}}
    """
    p = Path(xlsx_path)
    if not p.exists():
        return {}

    sheet = "EN" if (lang or "CN").strip().upper() == "EN" else "CN"
    try:
        raw = pd.read_excel(p, sheet_name=sheet, header=None)
    except Exception:
        raw = pd.read_excel(p, sheet_name=0, header=None)
    if raw.shape[0] < 8 or raw.shape[1] < 2:
        return {}

    header_dir = raw.iloc[0].ffill()
    header_scen = raw.iloc[1].fillna("")
    spec_blocks = _detect_capsule_spec_blocks(raw)
    if not spec_blocks:
        return {}

    out: Dict[str, Dict[str, Dict[str, object]]] = {}
    for col_idx in range(1, raw.shape[1]):
        direction = _normalize_text(header_dir[col_idx])
        scenario = _normalize_text(header_scen[col_idx])
        if not direction or not scenario:
            continue

        specs: List[Dict[str, str]] = []
        for start_row, spec_label in spec_blocks:
            clinical = _normalize_text(raw.iloc[start_row, col_idx])
            excipients = _normalize_text(raw.iloc[start_row + 1, col_idx])
            total = _normalize_text(raw.iloc[start_row + 2, col_idx])
            if not (clinical or excipients or total):
                continue
            specs.append(
                {
                    "spec": spec_label,
                    "clinical": clinical,
                    "excipients": excipients,
                    "total": total,
                }
            )

        out.setdefault(direction, {})[scenario] = {
            "scenario": scenario,
            "direction": direction,
            "specs": specs,
        }

    return out


def _pick_capsule_scenario(query: str, candidates: List[str]) -> str:
    q = (query or "").strip()
    if not q or not candidates:
        return ""

    # 关键词规则优先（处理缩写/命名差异）
    rules = [
        (("细菌性阴道炎", "BV"), ("BV",)),
        (("真菌性阴道炎", "霉菌", "CCV"), ("CCV", "真菌")),
        (("妊娠", "糖代谢", "GDM"), ("GDM", "糖", "血糖")),
        (("呼吸道", "肺炎"), ("肺炎", "呼吸")),
    ]
    for triggers, targets in rules:
        if any(t in q for t in triggers):
            for cand in candidates:
                if any(k in cand for k in targets):
                    return cand

    return _best_title_match(q, candidates)


def _parse_capsule_clinical(text: str) -> Tuple[str, str]:
    s = (text or "").strip().replace("：", ":")
    if not s:
        return "", ""
    if ":" in s:
        left, right = s.split(":", 1)
        return left.strip(), right.strip()
    return s, ""


def _excipient_name_only(item: str) -> str:
    s = (item or "").strip()
    if not s:
        return ""
    s = re.sub(r"^[•\-\s]+", "", s).strip()
    parts = re.split(r"[:：]", s, maxsplit=1)
    name = (parts[0] or "").strip()
    name = re.sub(r"\s+", " ", name).strip()
    return name


def _is_filler_excipient(name: str, lang: str) -> bool:
    s = (name or "").strip()
    if not s:
        return False
    is_en = (lang or "CN").strip().upper() == "EN"
    if is_en:
        key = re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()
        fillers = {
            "gum arabic",
            "arabic gum",
            "inulin",
            "fructooligosaccharides",
            "fructo oligosaccharides",
            "fos",
            "maltodextrin",
            "malto dextrin",
            "potato starch",
            "starch",
            "silicon dioxide",
            "magnesium stearate",
        }
        if key in fillers:
            return True
        # Handle combined labels like "Fructooligosaccharides (FOS)".
        if re.search(r"\binulin\b", key):
            return True
        if re.search(r"\bfos\b", key):
            return True
        if "fructooligosaccharides" in key or "fructo oligosaccharides" in key:
            return True
        return False
    fillers_cn = (
        "阿拉伯胶",
        "马铃薯淀粉",
        "二氧化硅",
        "硬脂酸镁",
        "淀粉",
        "麦芽糊精",
        "菊粉",
        "低聚果糖",
        "果寡糖",
    )
    return any(f in s for f in fillers_cn)


def _split_capsule_excipients(text: str) -> List[str]:
    s = (text or "").strip()
    if not s:
        return []
    parts = [p.strip() for p in re.split(r"[，,;；]+", s) if p.strip()]
    return parts


def _format_capsule_excipient_item(item: str, lang: str) -> str:
    s = (item or "").strip()
    if not s:
        return ""
    if (lang or "CN").strip().upper() != "EN":
        return s

    m = re.match(
        r"^(?P<name>.+?)\s+(?P<amount>\d+(?:\.\d+)?)\s*(?P<unit>mg|µg|ug|g)\b",
        s,
        flags=re.IGNORECASE,
    )
    if not m:
        return s

    name = m.group("name").strip().rstrip(":：")
    amount = m.group("amount").strip()
    unit_raw = m.group("unit").strip()
    unit = "µg" if unit_raw.lower() == "ug" else unit_raw.lower()
    unit = "µg" if unit == "µg" else unit
    return f"{name}: {amount}{unit}"


def _strip_mass_units(text: str) -> str:
    """移除形如 25mg/5.5mg/40 µg 的单位，仅保留数字（用于规格模块统一单位展示）。"""
    s = (text or "").strip()
    if not s:
        return ""
    return re.sub(
        r"(?P<num>\d+(?:\.\d+)?)\s*(?:mg|µg|ug|g)\b",
        r"\g<num>",
        s,
        flags=re.IGNORECASE,
    ).strip()


def _normalize_total_text(text: str) -> str:
    s = (text or "").strip().replace("：", ":")
    if not s:
        return ""
    m = re.match(r"(?i)^total\s*:\s*(.+)$", s)
    if m:
        return m.group(1).strip()
    return s


def _italicize_microbe_tokens_markdown(text: str) -> str:
    s = str(text or "")
    if not s:
        return ""
    # Ensure abbreviated species names are italic in Markdown.
    s = re.sub(r"\bH\.\s*pylori\b", "*H. pylori*", s, flags=re.IGNORECASE)
    return s


def _italicize_microbe_tokens_html(text: str) -> str:
    s = html.escape(str(text or ""))
    if not s:
        return ""
    # Ensure abbreviated species names are italic in HTML.
    s = re.sub(
        r"\bH\.\s*pylori\b",
        "<span class='latin'>H. pylori</span>",
        s,
        flags=re.IGNORECASE,
    )
    return s


@st.cache_data
def load_pdf_bytes(pdf_path: str, _cache_buster: float | None = None) -> bytes:
    p = Path(pdf_path)
    if not p.exists():
        return b""
    return p.read_bytes()


def _safe_filename_component(text: str) -> str:
    name = (text or "").strip()
    if not name:
        return "solution"
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120] if name else "solution"


def _ensure_wecpro_registered(text: str) -> str:
    s = (text or "").strip()
    if not s:
        return ""
    return re.sub(r"WecPro(?!®)", "WecPro®", s)


def _format_tm_sup_html(text: str, add_if_missing: bool = False) -> str:
    """将商品名中的 TM/™ 渲染为上标；可选在缺失时补 TM。"""
    s = (text or "").strip()
    if not s:
        return ""
    escaped = html.escape(s)
    escaped = re.sub(r"(?i)\bTM\b", "<sup>TM</sup>", escaped)
    escaped = escaped.replace("™", "<sup>TM</sup>")
    if add_if_missing and "<sup>TM</sup>" not in escaped:
        escaped = f"{escaped}<sup>TM</sup>"
    return escaped


def _extract_strain_codes(text: str) -> List[str]:
    """从文本中提取菌株代号（如 LRa05 / BLa80），保持出现顺序并去重。"""
    out: List[str] = []
    for m in re.finditer(r"([A-Za-z]{1,6}\d{2,3})", (text or "").strip()):
        code = m.group(1)
        if code not in out:
            out.append(code)
    return out


def _format_sci_name_markdown(sci: str) -> str:
    """Markdown 形式：种属名斜体，但 subsp. 不斜体。"""
    s = (sci or "").strip()
    if not s:
        return ""
    m = re.search(r"\bsubsp\.?\s+(?P<right>.+)$", s, flags=re.IGNORECASE)
    if m:
        left = s[: m.start()].strip()
        right = (m.group("right") or "").strip()
        if left and right:
            return f"*{left}* subsp. *{right}*"
    return f"*{s}*"


def _format_sci_name_html(sci: str) -> str:
    """HTML 形式：种属名斜体，但 subsp. 不斜体。"""
    s = (sci or "").strip()
    if not s:
        return ""
    m = re.search(r"\bsubsp\.?\s+(?P<right>.+)$", s, flags=re.IGNORECASE)
    if m:
        left = s[: m.start()].strip()
        right = (m.group("right") or "").strip()
        if left and right:
            return (
                f"<span class='latin'>{html.escape(left)}</span> "
                "<span class='latin-noi'>subsp.</span> "
                f"<span class='latin'>{html.escape(right)}</span>"
            )
    return f"<span class='latin'>{html.escape(s)}</span>"


def _to_english_formula(text: str) -> str:
    """把中文“核心配方”字符串尽量转换为英文（基于代号映射）。"""
    codes = _extract_strain_codes(text)
    if not codes:
        return (text or "").strip()

    parts: List[str] = []
    for code in codes:
        lookup = code
        prefix = ""
        display_code = code
        # pasteurized Akkermansia in Solutions decks often appears as pAkk11
        if code.startswith("pAkk") and code[1:] in _STRAIN_SCI_NAMES:
            prefix = "pasteurized "
            lookup = code[1:]
            display_code = code[1:]

        sci = _STRAIN_SCI_NAMES.get(lookup)
        if sci:
            sci_md = _format_sci_name_markdown(sci)
            parts.append(f"{prefix}{sci_md} {display_code}")
        else:
            parts.append(display_code)

    return ", ".join(parts)


_FORMULA_ITEM_SPLIT_RE = re.compile(r"[、，,;；]+")


def _split_formula_items(text: str) -> List[str]:
    s = (text or "").strip()
    if not s:
        return []
    return [p.strip() for p in _FORMULA_ITEM_SPLIT_RE.split(s) if p.strip()]


def _to_english_formula_html_items(text: str) -> List[str]:
    """Return EN core-formula items as HTML fragments (with italic scientific names)."""
    codes = _extract_strain_codes(text)
    if not codes:
        return [_italicize_microbe_tokens_html(x) for x in _split_formula_items(text)]

    out: List[str] = []
    for code in codes:
        lookup = code
        prefix = ""
        display_code = code
        if code.startswith("pAkk") and code[1:] in _STRAIN_SCI_NAMES:
            prefix = "pasteurized "
            lookup = code[1:]
            display_code = code[1:]

        sci = _STRAIN_SCI_NAMES.get(lookup)
        if sci:
            out.append(
                f"{html.escape(prefix)}{_format_sci_name_html(sci)} "
                f"<span class='formula-code'>{html.escape(display_code)}</span>"
            )
        else:
            out.append(f"<span class='formula-code'>{html.escape(display_code)}</span>")
    return out


def _colorize_solution_formula_html(formula_text: str, ui_lang: str) -> str:
    """Render Core Formula with colorful strain text for client-facing readability."""
    s = (formula_text or "").strip()
    if not s:
        return ""

    lang_norm = (ui_lang or "CN").strip().upper()
    if lang_norm == "EN":
        items_html = _to_english_formula_html_items(s)
        sep = ", "
    else:
        items_html = [_italicize_microbe_tokens_html(x) for x in _split_formula_items(s)]
        sep = "、"

    if not items_html:
        return _italicize_microbe_tokens_html(s)

    sep_html = f"<span class='formula-sep'>{html.escape(sep)}</span>"
    colored = [
        f"<span class='formula-strain formula-strain-{(i % 6) + 1}'>{item}</span>"
        for i, item in enumerate(items_html)
    ]
    return sep_html.join(colored)


def _parse_trial_entries(trial_lines: List[str]) -> List[Tuple[str, List[str]]]:
    """把 PPT 中的临床研究段落解析成 [(菌株/组合, [NCT/ChiCTR...]), ...]。"""
    out: List[Tuple[str, List[str]]] = []
    current_key = ""

    def extract_ids(text: str) -> List[str]:
        return re.findall(r"(NCT\d+|ChiCTR\d+)", text or "")

    for raw in trial_lines or []:
        line = str(raw).strip()
        if not line:
            continue

        m = re.match(r"^([A-Za-z0-9+/_-]+)\s*[:：]\s*(.*)$", line)
        if m:
            key = m.group(1).strip()
            rest = m.group(2).strip()
            current_key = key
            ids = extract_ids(rest)
            out.append((key, list(dict.fromkeys(ids))))
            continue

        ids = extract_ids(line)
        if ids and out and current_key:
            last_key, last_ids = out[-1]
            if last_key == current_key:
                for one in ids:
                    if one not in last_ids:
                        last_ids.append(one)
            continue

    # 去掉空 key / 空 id
    cleaned: List[Tuple[str, List[str]]] = []
    for key, ids in out:
        k = (key or "").strip()
        if not k:
            continue
        uniq = [x for x in dict.fromkeys([i for i in ids if i])]
        if not uniq:
            continue
        cleaned.append((k, uniq))
    return cleaned


@st.cache_data
def build_solution_pdf_bytes(
    source_pdf_path: str,
    start_page: int,
    end_page: int,
    _cache_buster: float | None = None,
) -> bytes:
    """从整本 PDF 中裁剪指定页范围（1-based, 含首尾），返回新 PDF bytes。"""
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore[import-not-found]
    except Exception:
        return b""

    start = max(1, int(start_page))
    end = max(start, int(end_page))

    reader = PdfReader(source_pdf_path)
    total = len(reader.pages)

    writer = PdfWriter()
    for page in range(start, end + 1):
        idx = page - 1
        if 0 <= idx < total:
            writer.add_page(reader.pages[idx])

    if len(writer.pages) == 0:
        return b""

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


@st.cache_data
def render_pdf_pages_png(
    pdf_path: str,
    pages: Tuple[int, ...],
    scale: float = 2.0,
    _cache_buster: float | None = None,
) -> List[bytes]:
    """把指定 PDF 页渲染成 PNG（用于网页内原版展示）。"""
    try:
        import fitz  # type: ignore[import-not-found]
    except Exception:
        return []

    p = Path(pdf_path)
    if not p.exists():
        return []

    try:
        s = float(scale) if scale else 2.0
    except Exception:
        s = 2.0
    s = min(max(s, 1.0), 3.0)

    images: List[bytes] = []
    doc = fitz.open(str(p))
    try:
        total = int(getattr(doc, "page_count", len(doc)))
        for page_no in pages:
            idx = max(1, int(page_no)) - 1
            if idx < 0 or idx >= total:
                continue
            page = doc.load_page(idx)
            pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
            images.append(pix.tobytes("png"))
    finally:
        doc.close()

    return images


def list_docs(category: str) -> List[Path]:
    """返回某功能方向下的说明书文件列表。"""
    cat_dir = DOCS_DIR / category
    if not cat_dir.exists():
        return []
    return [p for p in cat_dir.iterdir() if p.is_file()]

_CATEGORY_THEME: Dict[str, Dict[str, str]] = {
    # Colors are matched to the 7 category styles used in the PDF deck.
    "女性健康": {"accent1": "#901050", "accent2": "#F472B6", "accent3": "#FB7185", "tint": "#FCE7F3"},
    "情绪健康": {"accent1": "#A02010", "accent2": "#FB923C", "accent3": "#F59E0B", "tint": "#FFEDD5"},
    "代谢健康": {"accent1": "#204020", "accent2": "#22C55E", "accent3": "#34D399", "tint": "#DCFCE7"},
    "胃肠健康": {"accent1": "#503010", "accent2": "#EAB308", "accent3": "#FBBF24", "tint": "#FEF3C7"},
    "免疫健康": {"accent1": "#003030", "accent2": "#14B8A6", "accent3": "#2DD4BF", "tint": "#CCFBF1"},
    "婴童健康": {"accent1": "#402060", "accent2": "#A78BFA", "accent3": "#C4B5FD", "tint": "#F3E8FF"},
    "口腔健康": {"accent1": "#003070", "accent2": "#60A5FA", "accent3": "#93C5FD", "tint": "#DBEAFE"},
}

_CATEGORY_LABELS_EN: Dict[str, str] = {
    "女性健康": "Women's Health",
    "情绪健康": "Mental Health",
    "代谢健康": "Metabolic Health",
    "胃肠健康": "Gastrointestinal Health",
    "免疫健康": "Immunological Health",
    "婴童健康": "Infant Health",
    "口腔健康": "Dental & Oral Health",
}

_WECPRO_FORMULA_BENEFIT_EN: Dict[str, str] = {
    "女性健康": "Supports vaginal microbiome balance, helps address vaginitis-related concerns, and promotes hormonal and metabolic homeostasis for women’s well-being.",
    "情绪健康": "Helps manage stress and mood, improves sleep quality, and supports relief of anxiety and depressive symptoms.",
    "代谢健康": "Supports metabolic balance and weight management, including healthier control of blood glucose, lipids, and blood pressure.",
    "胃肠健康": "Supports gastrointestinal function and microbiome balance; helps relieve constipation and diarrhea; supports gut motility and intestinal barrier health.",
    "免疫健康": "Supports immune defenses, helps reduce allergic responses and related inflammation, and promotes immune homeostasis.",
    "婴童健康": "Supports early-life development and the establishment of immune and gut microbiome homeostasis.",
    "口腔健康": "Supports oral microbiome balance and local immunity, promotes periodontal health, and helps maintain fresh breath.",
}

_WECPRO_FORMULA_VARIANTS: Dict[str, List[Dict[str, object]]] = {
    "胃肠健康": [
        {
            "tag": {"CN": "高端款", "EN": "Premium"},
            "product": {"CN": "WecPro®-GIHealth805", "EN": "WecPro®-GIHealth805"},
            "benefit": {
                "CN": "调节胃肠健康，改善便秘与腹泻，支持胃肠运动功能与菌群稳态，缓解肠道损伤",
                "EN": "Supports gastrointestinal health, helps relieve constipation and diarrhea, supports gut motility and microbiome homeostasis, and helps ease intestinal injury.",
            },
            "core_cn": "动物双歧杆菌乳亚种BLa80、鼠李糖乳酪杆菌LRa05",
            "codes": ["BLa80", "LRa05"],
        },
	        {
	            "tag": {"CN": "基础款", "EN": "Base"},
	            "product": {"CN": "WecPro®-GUT99", "EN": "WecPro®-GUT99"},
	            "benefit": {
	                "CN": "作为肠道健康基础配方，重建菌群稳态并支持肠道蠕动与屏障功能，帮助维持长期消化舒适与排便规律",
	                "EN": "A foundational gut-health formula to rebuild microbiome homeostasis and support gut motility and barrier function for long-term digestive comfort and regularity.",
	            },
	            "core_cn": "动物双歧杆菌乳亚种BLa36、乳酸片球菌PA53、植物乳植杆菌Lp18、凝结魏茨曼氏菌BC99",
	            "codes": ["BLa36", "PA53", "Lp18", "BC99"],
	        },
        {
            "tag": {"CN": "高活性益生菌酸奶款", "EN": "Active Probiotic Yogurt"},
            "product": {"CN": "WecPro®-DigestBi", "EN": "WecPro®-DigestBi"},
            "benefit": {
                "CN": "高活性益生菌发酵酸奶，支持肠道蠕动与菌群稳态，改善便秘相关不适，打造日常肠道舒适底盘",
                "EN": "A high-activity probiotic fermented yogurt concept that supports gut motility and microbiome balance, helping relieve constipation-related discomfort for daily gut comfort.",
            },
            "core_cn": "动物双歧杆菌乳亚种BLa80、长双歧杆菌长亚种BL21、短双歧杆菌BBr60、青春双歧杆菌BAC30、长双歧杆菌婴儿亚种BI45",
            "codes": ["BLa80", "BL21", "BBr60", "BAC30", "BI45"],
        },
    ]
}

_SERIES_OPTIONS = ["WecLac", "WecPro® Formula", "WecPro® Solution"]

_SERIES_THEME: Dict[str, Dict[str, str]] = {
    "WecLac": {"accent1": "#7C3AED", "accent2": "#FF2D55", "accent3": "#0A84FF", "tint": "#FFF1F2"},
    "WecPro® Formula": {"accent1": "#4F46E5", "accent2": "#EC4899", "accent3": "#22C55E", "tint": "#EEF2FF"},
    "WecPro® Solution": {"accent1": "#6366F1", "accent2": "#EC4899", "accent3": "#22C55E", "tint": "#F8FAFC"},
}


def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = (hex_color or "").strip().lstrip("#")
    if len(h) == 3:
        h = "".join([c * 2 for c in h])
    if len(h) != 6:
        return (99, 102, 241)  # indigo fallback
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _rgba(hex_color: str, alpha: float) -> str:
    r, g, b = _hex_to_rgb(hex_color)
    try:
        a = float(alpha)
    except Exception:
        a = 1.0
    a = min(max(a, 0.0), 1.0)
    return f"rgba({r},{g},{b},{a})"


def _render_header(series: str = "", category: str = "", badge: str = "") -> None:
    series = (series or "").strip()
    category = (category or "").strip()
    badge = (badge or "").strip()
    ui_lang = str(st.session_state.get("ui_lang", "CN")).strip().upper() or "CN"

    theme: Dict[str, str] = {}
    if category and category in _CATEGORY_THEME:
        theme = _CATEGORY_THEME[category]
    elif series and series in _SERIES_THEME:
        theme = _SERIES_THEME[series]

    accent1 = theme.get("accent1", "#6366f1")
    accent2 = theme.get("accent2", "#ec4899")
    accent3 = theme.get("accent3", "#22c55e")
    tint = theme.get("tint", "rgba(255,255,255,0.75)")
    r1, g1, b1 = _hex_to_rgb(accent1)
    r2, g2, b2 = _hex_to_rgb(accent2)
    r3, g3, b3 = _hex_to_rgb(accent3)

    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800;900&display=swap');
        /* 更适合对外展示的商务风样式（干净克制、更像企业官网） */
        :root{
          --bg: #f4f6fb;
          --card: rgba(255,255,255,0.96);
          --border: rgba(15,23,42,0.14);
          --shadow: 0 12px 34px rgba(2,6,23,0.06);
          --text: #0f172a;
          --muted: #334155;
          --accent1: #1d4ed8;
          --accent2: #0ea5e9;
          --accent3: #10b981;
          --tint: rgba(255,255,255,0.78);
          --accent1-rgb: 29,78,216;
          --accent2-rgb: 14,165,233;
          --accent3-rgb: 16,185,129;
          /* Fixed brand identity — never overridden by series theme */
          --brand: #D10025;
          --brand-rgb: 209,0,37;
          --brand-dark: #a8001e;
        }

        html, body, [class*="css"]  {
          font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                       "Helvetica Neue", Arial, "Noto Sans", "Apple Color Emoji", "Segoe UI Emoji";
          color: var(--text);
        }

        [data-testid="stAppViewContainer"]{
          background:
            radial-gradient(980px 520px at 18% 10%, rgba(var(--accent1-rgb),0.08), transparent 60%),
            radial-gradient(860px 520px at 86% 8%, rgba(var(--accent2-rgb),0.06), transparent 60%),
            linear-gradient(180deg, rgba(15,23,42,0.03), transparent 40%),
            var(--bg);
        }

        [data-testid="stHeaderActionElements"] { display: none; }
        /* Remove Streamlit top “notch” header bar */
        header[data-testid="stHeader"]{ display: none; }
        [data-testid="stToolbar"]{ display: none; }
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }

        .block-container { padding-top: 1.1rem; padding-bottom: 2.5rem; max-width: 1180px; }

        /* Streamlit dialogs (make larger & more client-friendly) */
        div[role="dialog"]{
          width: min(1080px, 96vw) !important;
          max-width: min(1080px, 96vw) !important;
        }
        div[role="dialog"] [data-testid="stMarkdownContainer"]{
          font-size: 0.95rem;
          line-height: 1.55;
        }

        /* Card-like containers (st.container(border=True)) */
        [data-testid="stVerticalBlockBorderWrapper"]{
          background: var(--card);
          border: 1px solid var(--border);
          border-radius: 18px;
          box-shadow: var(--shadow);
          position: relative;
          z-index: 0; /* create stacking context for watermark */
        }
        /* WecLac cards are rendered via HTML (no border=True container). */

        /* WecLac: click directly on IP image (no Streamlit button) */
        .weclac-open{
          display: inline-block;
          line-height: 0;
          border-radius: 16px;
          overflow: hidden;
          cursor: pointer;
          text-decoration: none;
        }
        .weclac-open:focus,
        .weclac-open:focus-visible{
          outline: none;
          box-shadow: 0 0 0 3px rgba(var(--accent1-rgb),0.25);
        }
        .weclac-open img{
          display: block;
        }
        [data-testid="stVerticalBlockBorderWrapper"]:has(.weclac-card-scope) .ip-card{
          box-shadow: none;
          background: transparent;
          border: 0;
          padding: 14px 14px 12px 14px;
        }

        /* Tabs */
        div[data-testid="stTabs"] button[data-baseweb="tab"]{
          border-radius: 999px;
          padding: 8px 14px;
          margin-right: 8px;
          background: rgba(255,255,255,0.65);
          border: 1px solid rgba(15,23,42,0.10);
        }
        div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"]{
          background: rgba(var(--accent1-rgb),0.12);
          border-color: rgba(var(--accent1-rgb),0.28);
          color: var(--text);
        }

        /* WecPro® Formula: variants grid (e.g., GI has 3 formulas) */
        .v-grid{
          display:grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 12px;
        }
        .v-box{
          border-radius: 18px;
          border: 1px solid rgba(15,23,42,0.10);
          background: rgba(255,255,255,0.90);
          padding: 12px 12px 10px;
        }
        .v-title{
          font-weight: 950;
          letter-spacing: -0.01em;
          line-height: 1.15;
          font-size: 0.98rem;
          display:flex;
          align-items:center;
          justify-content:space-between;
          gap: 8px;
        }
        .v-tag{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          padding: 4px 8px;
          border-radius: 999px;
          border: 1px solid rgba(var(--accent1-rgb),0.22);
          background: rgba(255,255,255,0.70);
          color: var(--muted);
          font-weight: 850;
          font-size: 0.78rem;
          white-space: nowrap;
        }
        .v-meta{
          color: var(--muted);
          font-size: 0.78rem;
          margin-top: 10px;
          font-weight: 800;
        }
        .v-text{
          font-size: 0.90rem;
          line-height: 1.45;
          color: rgba(15,23,42,0.84);
        }

        /* Hero top controls: same-page segmented controls (no jump links) */
        .hero-seg-row{
          margin-top: 16px;
        }
        [data-testid="stSegmentedControl"]{
          background: rgba(255,255,255,0.82) !important;
          border: 1px solid rgba(15,23,42,0.12) !important;
          border-radius: 999px !important;
          padding: 4px !important;
          box-shadow: 0 8px 20px rgba(15,23,42,0.06) !important;
        }
        [data-testid="stSegmentedControl"] button{
          min-height: 37px !important;
          border-radius: 999px !important;
          border: 1px solid transparent !important;
          font-size: 0.88rem !important;
          font-weight: 780 !important;
          letter-spacing: 0.01em !important;
          color: rgba(15,23,42,0.66) !important;
          background: transparent !important;
          box-shadow: none !important;
          transition: all .18s ease;
        }
        [data-testid="stSegmentedControl"] button:hover{
          background: rgba(var(--accent1-rgb),0.10) !important;
          color: rgba(15,23,42,0.92) !important;
        }
        [data-testid="stSegmentedControl"] button[aria-selected="true"],
        [data-testid="stSegmentedControl"] button[aria-pressed="true"],
        [data-testid="stSegmentedControl"] button[aria-checked="true"]{
          background: linear-gradient(135deg, rgba(var(--accent1-rgb),0.94), rgba(var(--accent2-rgb),0.90)) !important;
          border-color: rgba(var(--accent1-rgb),0.55) !important;
          color: #fff !important;
          box-shadow: 0 10px 24px rgba(var(--accent1-rgb),0.22) !important;
        }
        .hero-seg-row + div[data-testid="stHorizontalBlock"]{
          align-items: center !important;
          gap: 0.9rem !important;
        }
        .hero-seg-row + div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:last-child{
          display: flex;
          justify-content: flex-end;
          min-width: 164px;
        }
        .hero-seg-row + div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:last-child [data-testid="stSegmentedControl"]{
          width: 164px !important;
          margin-left: auto;
        }

        /* Download button */
        [data-testid="stDownloadButton"] button{
          background: var(--accent1);
          border: 1px solid rgba(15,23,42,0.10);
        }
        [data-testid="stDownloadButton"] button p{ color: #fff; font-weight: 600; }

        /* ── Hero: split layout (text + visual) ── */
        .hero-wrap{
          position: relative;
          border-radius: 22px;
          min-height: 0;
          display: grid;
          grid-template-columns: minmax(0, 1.08fr) minmax(320px, 0.92fr);
          gap: 14px;
          align-items: stretch;
          padding: 14px 16px 14px 10px;
          background:
            radial-gradient(760px 360px at 16% 10%, rgba(var(--accent1-rgb),0.07), transparent 62%),
            linear-gradient(145deg, rgba(255,255,255,0.88), rgba(255,255,255,0.78));
        }
        .hero-content{
          position: relative;
          z-index: 1;
          padding: 2px 4px 2px 0;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          min-width: 0;
        }
        /* ── Right visual panel: 5-factory infographic ── */
        .hero-visual{
          position: relative;
          border-radius: 16px;
          overflow: hidden;
          border: 1px solid rgba(15,23,42,0.10);
          box-shadow: 0 10px 26px rgba(2,6,23,0.09);
          min-height: 186px;
          background: linear-gradient(135deg, rgba(15,23,42,0.05), rgba(15,23,42,0.02));
        }
        .hero-photo{
          position: absolute;
          inset: 0;
          width: 100%;
          height: 100%;
          object-fit: cover;
          object-position: center center;
          filter: saturate(0.95) contrast(1.02) brightness(0.99);
          display: block;
        }
        .hero-visual-wm{
          display: none;
        }
        .hero-visual-meta{
          display: none;
        }
        .hero-visual-chip{
          display: none;
        }
        /* ── Logo lockup: free-standing W monogram + wordmark ── */
        .hero-logo{
          display: inline-flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 14px;
          margin-left: 0;
          line-height: 1;
        }
        /* The W SVG path rendered large, no container */
        .hero-logo-w-svg{
          display: block;
          width: 52px;
          height: auto;
          flex: 0 0 auto;
          filter: drop-shadow(0 2px 8px rgba(209,0,37,0.22));
        }
        /* Thin vertical hairline divider */
        .hero-logo-divider{
          width: 1px;
          height: 32px;
          background: rgba(15,23,42,0.14);
          flex: 0 0 auto;
          margin: 0;
        }
        /* Wordmark: stacked two-line */
        .hero-logo-text{
          display: flex;
          flex-direction: column;
          gap: 4px;
          margin-left: -2px;
          line-height: 1;
        }
        /* Wordmark SVG */
        .hero-logo-wordmark{
          display: inline-block;
          flex: 0 0 auto;
          font-family: "Avenir Next Rounded", "Nunito", "Manrope", "Avenir Next", "Segoe UI", sans-serif;
          font-size: 2.02rem;
          font-weight: 800;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          line-height: 0.95;
          background: linear-gradient(92deg, #0f172a 0%, #1e293b 48%, #0f4c81 100%);
          -webkit-background-clip: text;
          background-clip: text;
          color: transparent;
          -webkit-text-fill-color: transparent;
          text-shadow: 0 0 16px rgba(56, 189, 248, 0.12);
        }
        .hero-logo-sub{
          font-size: 0.50rem;
          font-weight: 650;
          letter-spacing: 0.30em;
          text-transform: uppercase;
          color: rgba(15,23,42,0.46);
          line-height: 1;
          margin-top: 2px;
          display: block;
        }
        /* ── Title with brand-red accent ── */
        .hero-title{
          font-size: 1.72rem;
          font-weight: 800;
          letter-spacing: -0.02em;
          line-height: 1.18;
          margin: 0;
          position: relative;
          padding-left: 14px;
        }
        .hero-title::before{
          content: "";
          position: absolute;
          left: 0;
          top: 0.12em;
          bottom: 0.12em;
          width: 3px;
          border-radius: 999px;
          background: var(--brand);
        }
        .hero-head{
          display: flex;
          flex-direction: column;
          gap: 0;
        }
        /* ── Hero card: brand-red left accent strip ── */
        [data-testid="stVerticalBlockBorderWrapper"]:first-of-type{
          border-left: 3px solid var(--brand) !important;
          border-radius: 20px !important;
          overflow: hidden !important;
        }
        /* Keep inner padding so nav controls stay inside card */
        [data-testid="stVerticalBlockBorderWrapper"]:first-of-type > div > div[data-testid="stVerticalBlock"] > div:first-child{
          padding: 0 !important;
        }
        /* Legacy hero-art classes: hidden */
        .hero-art{ display: none; }
        .hero-art-photo{ display: none; }
        .hero-art-wm{ display: none; }
        @media (max-width: 720px){
          .hero-wrap{
            grid-template-columns: 1fr;
            min-height: 0;
            padding: 14px;
            gap: 12px;
          }
          .hero-visual{
            min-height: 164px;
          }
        }
        .hero-mark{
          width: 44px;
          height: 44px;
          margin-top: 3px;
          flex: 0 0 auto;
          background: linear-gradient(135deg, rgba(15,23,42,0.88), rgba(100,116,139,0.88));
          opacity: 0.95;
          -webkit-mask-repeat: no-repeat;
          -webkit-mask-position: center;
          -webkit-mask-size: contain;
          mask-repeat: no-repeat;
          mask-position: center;
          mask-size: contain;
          filter: drop-shadow(0 10px 22px rgba(2,6,23,0.08));
        }
        .hero-subtitle{
          margin-top: 0.25rem;
          color: var(--muted);
          font-size: 0.95rem;
        }
        .hero-desc{
          margin-top: 0.55rem;
          color: var(--muted);
          font-size: 0.98rem;
          line-height: 1.5;
        }
        .hero-desc strong{
          color: var(--brand);
          font-weight: 750;
        }
        .hero-series-label{
          margin-top: 0.75rem;
          color: var(--muted);
          font-size: 0.86rem;
          font-weight: 650;
          letter-spacing: 0.02em;
          text-transform: uppercase;
        }
        .pill{
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 6px 12px;
          border-radius: 999px;
          border: 1px solid var(--border);
          background: rgba(255,255,255,0.75);
          font-weight: 600;
        }
        .core-formula-line{
          font-size: 1.02rem;
          line-height: 1.72;
          color: var(--text);
        }
        .core-formula-name{
          font-weight: 920;
          color: var(--text);
        }
        .core-formula-sep{
          color: rgba(15,23,42,0.68);
        }
        .formula-strain{
          font-weight: 820;
          letter-spacing: 0.01em;
        }
        .formula-strain .latin,
        .formula-strain .latin-noi,
        .formula-strain .formula-code{
          color: inherit !important;
        }
        .formula-strain-1{ color: var(--accent1); }
        .formula-strain-2{ color: var(--accent2); }
        .formula-strain-3{ color: #0EA5E9; }
        .formula-strain-4{ color: #16A34A; }
        .formula-strain-5{ color: #D97706; }
        .formula-strain-6{ color: #7C3AED; }
        .formula-sep{
          color: rgba(15,23,42,0.45);
          padding: 0 2px;
        }
        .formula-code{
          color: rgba(15,23,42,0.92);
          font-weight: 900;
        }
        .core-func-title{
          margin-top: 10px;
          margin-bottom: 3px;
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 5px 12px;
          border-radius: 999px;
          border: 1px solid rgba(var(--accent1-rgb),0.24);
          background: linear-gradient(
            90deg,
            rgba(var(--accent1-rgb),0.09),
            rgba(var(--accent2-rgb),0.09)
          );
          color: var(--text);
          font-weight: 900;
          letter-spacing: 0.01em;
        }
        .core-func-dot{
          width: 8px;
          height: 8px;
          border-radius: 999px;
          background: linear-gradient(90deg, var(--accent1), var(--accent2));
          box-shadow: 0 0 0 3px rgba(var(--accent2-rgb),0.16);
        }
        .core-func-list{
          margin: 2px 0 0 0 !important;
          padding-left: 1.2rem;
          line-height: 1.52;
        }
        .core-func-list li{
          margin: 0.22rem 0;
          color: var(--text);
          font-weight: 620;
        }

        .hero-badge{
          display: inline-flex;
          align-items: center;
          justify-content: center;
          padding: 8px 12px;
          border-radius: 999px;
          font-weight: 800;
          letter-spacing: -0.01em;
          color: #fff;
          background: var(--accent1);
          box-shadow: 0 10px 24px rgba(2,6,23,0.08);
          border: 1px solid rgba(255,255,255,0.25);
        }

        .spec-box{
          border-radius: 18px;
          border: 1px solid rgba(15,23,42,0.10);
          background: rgba(255,255,255,0.90);
          backdrop-filter: none;
          padding: 12px 12px 10px 12px;
          height: 100%;
          display:flex;
          flex-direction:column;
        }
        .spec-grid{
          display:grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 12px;
          margin-top: 10px;
          margin-bottom: 8px;
        }
        .spec-title{
          font-weight: 900;
          letter-spacing: -0.01em;
          margin: 0 0 6px 0;
          font-size: 1.02rem;
          line-height: 1.15;
        }
        .spec-meta{
          color: var(--muted);
          font-size: 0.82rem;
          margin-top: 10px;
        }
        .spec-list{
          display: flex;
          flex-direction: column;
          gap: 10px;
          margin-top: 6px;
        }
        .spec-line{
          display: flex;
          align-items: center;
          flex-wrap: wrap;
          gap: 8px;
          line-height: 1.45;
        }
        .spec-k{
          color: var(--text);
          font-weight: 860;
          letter-spacing: -0.005em;
        }
        .spec-v{
          color: rgba(15,23,42,0.86);
          font-weight: 660;
        }
        .spec-v-formula{
          display: inline-flex;
          align-items: center;
          flex-wrap: wrap;
          padding: 4px 12px;
          border-radius: 999px;
          border: 1px solid rgba(var(--accent1-rgb),0.22);
          background: linear-gradient(
            90deg,
            rgba(var(--accent1-rgb),0.10),
            rgba(var(--accent2-rgb),0.12)
          );
          color: var(--text);
          font-weight: 820;
          font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace;
          letter-spacing: 0.01em;
        }
        .spec-product-name{
          color: var(--accent1);
          font-weight: 920;
        }
        .spec-code{
          color: var(--accent2);
          font-weight: 900;
          letter-spacing: 0.01em;
        }
        .spec-plus{
          display: inline-block;
          margin: 0 8px;
          color: rgba(15,23,42,0.58);
          font-weight: 860;
        }
        .spec-checklist{
          display: inline-flex;
          align-items: center;
          flex-wrap: wrap;
          gap: 8px;
        }
        .spec-check-item{
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 4px 10px 4px 7px;
          border-radius: 999px;
          border: 1px solid rgba(var(--accent2-rgb),0.28);
          background: linear-gradient(
            90deg,
            rgba(var(--accent1-rgb),0.08),
            rgba(var(--accent2-rgb),0.10)
          );
          color: var(--text);
          font-weight: 760;
          letter-spacing: 0.005em;
        }
        .spec-check-dot{
          width: 18px;
          height: 18px;
          border-radius: 999px;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          background: linear-gradient(135deg, var(--accent1), var(--accent2));
          color: #fff;
          font-size: 0.74rem;
          font-weight: 920;
          box-shadow: 0 3px 8px rgba(var(--accent2-rgb),0.26);
          line-height: 1;
        }

        /* Generic tiles (for WecLac / Formula cards) */
        .tile-wrap{
          border-radius: 18px;
          padding: 6px;
          background: rgba(255,255,255,0.90);
          border: 1px solid rgba(var(--accent1-rgb),0.22);
          box-shadow: 0 14px 40px rgba(2,6,23,0.08);
        }
        .tile{
          border-radius: 14px;
          padding: 14px 14px 12px 14px;
          background: rgba(255,255,255,0.92);
          border: 1px solid rgba(15,23,42,0.08);
        }
        .tile-top{
          display:flex;
          align-items:flex-start;
          justify-content:space-between;
          gap: 10px;
        }
        .tile-title{
          font-weight: 900;
          letter-spacing: -0.01em;
          font-size: 1.05rem;
          line-height: 1.2;
        }
        .tile-badge{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          padding: 6px 10px;
          border-radius: 999px;
          background: rgba(2,6,23,0.06);
          border: 1px solid rgba(15,23,42,0.08);
          font-weight: 800;
          font-size: 0.82rem;
          white-space: nowrap;
        }
        a.tile-badge.tile-badge-link{
          text-decoration: none;
          color: var(--text);
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.86), rgba(255,255,255,0.86)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.75), rgba(var(--accent2-rgb),0.75)) border-box;
          box-shadow: 0 10px 22px rgba(2,6,23,0.08);
          transition: transform 120ms ease, box-shadow 120ms ease;
        }
        a.tile-badge.tile-badge-link:hover{
          transform: translateY(-1px);
          box-shadow: 0 12px 26px rgba(2,6,23,0.10);
        }
        .tile-badge-strong{
          background: rgba(255,255,255,0.7);
          border-color: rgba(255,255,255,0.45);
          color: var(--accent1);
        }
        .tile-desc{
          color: var(--muted);
          margin-top: 6px;
          line-height: 1.45;
          font-size: 0.93rem;
        }
        .tile-grid{
          display:grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 10px;
          margin-top: 10px;
        }
        .tile-metric{
          border-radius: 12px;
          padding: 8px 10px;
          background: rgba(15,23,42,0.04);
          border: 1px solid rgba(15,23,42,0.06);
        }
        .tile-k{
          color: var(--muted);
          font-weight: 800;
          font-size: 0.75rem;
        }
        .tile-v{
          font-weight: 900;
          font-size: 0.98rem;
          margin-top: 2px;
        }

        /* WecLac: IP grid + glass details */
        .ip-wrap{
          border-radius: 22px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.72), rgba(255,255,255,0.66)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.16), rgba(var(--accent2-rgb),0.16)) border-box;
          box-shadow: 0 14px 40px rgba(2,6,23,0.08);
          margin: 0 0 18px 0;
          height: auto;
        }
        .ip-card{
          border-radius: 22px;
          padding: 12px 12px 10px 12px;
          background: rgba(255,255,255,0.56);
          backdrop-filter: blur(12px);
          box-shadow: 0 14px 40px rgba(2,6,23,0.08);
          border: 0;
          min-height: 300px;
          margin: 0 0 16px 0;
          display:flex;
          flex-direction:column;
        }
        .ip-link{ text-decoration:none; display:block; }
        .ip-avatar{
          width: 96px;
          height: 96px;
          border-radius: 24px;
          overflow: hidden;
          margin: 0 auto;
          display:flex;
          align-items:center;
          justify-content:center;
          box-sizing: border-box;
          padding: 6px;
          background:
            linear-gradient(135deg, rgba(255,255,255,0.78), rgba(255,255,255,0.35));
          border: 1px solid rgba(255,255,255,0.65);
          box-shadow: 0 10px 26px rgba(2,6,23,0.10);
        }
        .ip-avatar img{
          width: 100%;
          height: 100%;
          object-fit: contain;
          transform: scale(1.16);
          transform-origin: center center;
        }
        .ip-code{
          margin-top: 10px;
          display:flex;
          flex-direction:column;
          justify-content:flex-start;
          align-items:center;
          gap: 6px;
          text-align:center;
          min-height: 48px;
        }
        .ip-latin{
          font-style: italic;
          font-weight: 900;
          font-size: 0.82rem;
          line-height: 1.2;
          color: rgba(15,23,42,0.80);
          max-width: 340px;
          white-space: normal;
          overflow: hidden;
          text-overflow: ellipsis;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
        }
        .code-pill{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          padding: 6px 10px;
          border-radius: 999px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.82)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.40), rgba(var(--accent2-rgb),0.40)) border-box;
          font-weight: 950;
          letter-spacing: 0.02em;
        }
        .ip-name{
          margin-top: 3px;
          color: var(--muted);
          text-align:center;
          font-size: 0.85rem;
          line-height: 1.25;
          min-height: 2.5em;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .ip-details{
          margin-top: auto;
          border-radius: 14px;
          background: rgba(255,255,255,0.40);
          border: 1px solid rgba(15,23,42,0.06);
          padding: 8px 10px;
        }
        .ip-details summary{
          cursor: pointer;
          user-select: none;
          font-weight: 850;
          color: var(--text);
          list-style: none;
        }
        .ip-details summary::-webkit-details-marker{ display:none; }
        .ip-details summary:after{
          content: "＋";
          float: right;
          color: rgba(15,23,42,0.55);
        }
        .ip-details[open] summary:after{ content:"－"; }
        .ip-kv{
          display:grid;
          grid-template-columns: minmax(84px, 108px) minmax(0, 1fr);
          gap: 6px 10px;
          margin-top: 8px;
          font-size: 0.82rem;
          line-height: 1.3;
        }
        .ip-k{
          color: var(--muted);
          font-weight: 850;
          white-space: normal;
          overflow-wrap: anywhere;
        }
        .ip-v{
          color: var(--text);
          font-weight: 650;
          min-width: 0;
          overflow-wrap: anywhere;
          word-break: break-word;
        }
        .chip{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          padding: 8px 12px;
          border-radius: 999px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.78), rgba(255,255,255,0.78)) padding-box,
            linear-gradient(90deg, var(--accent1), var(--accent2)) border-box;
          box-shadow: 0 10px 28px rgba(2,6,23,0.08);
          font-weight: 900;
          color: var(--text);
          white-space: nowrap;
        }
        .back-link{
          display:inline-flex;
          align-items:center;
          gap: 8px;
          padding: 8px 12px;
          border-radius: 999px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.78), rgba(255,255,255,0.78)) padding-box,
            linear-gradient(90deg, var(--accent1), var(--accent2)) border-box;
          text-decoration:none;
          color: var(--text);
          font-weight: 900;
        }
        .detail-wrap{
          border-radius: 22px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.72), rgba(255,255,255,0.66)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.24), rgba(var(--accent2-rgb),0.24)) border-box;
          box-shadow: 0 16px 48px rgba(2,6,23,0.08);
        }
        .detail-card{
          border-radius: 22px;
          padding: 18px;
          background: rgba(255,255,255,0.56);
          backdrop-filter: blur(12px);
        }
        .latin{
          font-style: italic;
          font-weight: 900;
          color: rgba(15,23,42,0.82);
        }
        .latin-noi{
          font-style: normal;
          font-weight: 900;
          color: rgba(15,23,42,0.82);
        }
        .detail-title{
          font-weight: 950;
          letter-spacing: -0.015em;
          font-size: 1.35rem;
          line-height: 1.2;
        }
        .detail-star{
          color: var(--accent2);
          margin-left: 8px;
        }
        .detail-grid{
          display:grid;
          grid-template-columns: 140px 1fr;
          gap: 16px;
          margin-top: 14px;
          align-items:start;
        }
        .detail-avatar{
          width: 140px;
          height: 140px;
          border-radius: 30px;
          overflow: hidden;
          display:flex;
          align-items:center;
          justify-content:center;
          background:
            linear-gradient(135deg, rgba(255,255,255,0.78), rgba(255,255,255,0.35));
          border: 1px solid rgba(255,255,255,0.70);
          box-shadow: 0 12px 30px rgba(2,6,23,0.10);
        }
        .detail-avatar img{ width:100%; height:100%; object-fit:contain; }
        .detail-sub{
          color: var(--muted);
          margin-top: 6px;
          line-height: 1.45;
          font-size: 0.95rem;
        }
        .kv-table{
          border-radius: 16px;
          background: rgba(15,23,42,0.04);
          border: 1px solid rgba(15,23,42,0.06);
          padding: 12px 12px;
        }
        .kv-grid{
          display:grid;
          grid-template-columns: 170px 1fr;
          gap: 8px 12px;
          font-size: 0.95rem;
          line-height: 1.35;
        }
        .kv-k{
          color: var(--muted);
          font-weight: 900;
          white-space: nowrap;
        }
        .kv-v{
          color: var(--text);
          font-weight: 650;
        }
        .clinical-grid{
          grid-template-columns: minmax(180px, 280px) minmax(0, 1fr);
          align-items: start;
        }
        .clinical-grid .kv-k{
          white-space: normal;
          overflow-wrap: anywhere;
          word-break: break-word;
        }
        .clinical-grid .kv-v{ min-width: 0; }

        /* WecPro® Formula: 7-row list */
        .f-row-wrap{
          border-radius: 22px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.72), rgba(255,255,255,0.66)) padding-box,
            linear-gradient(90deg, var(--row1), var(--row2)) border-box;
          box-shadow: 0 14px 40px rgba(2,6,23,0.08);
          margin: 0 0 14px 0;
          transition: transform 140ms ease, box-shadow 140ms ease;
        }
        .f-row-wrap:hover{
          transform: translateY(-1px);
          box-shadow: 0 16px 48px rgba(2,6,23,0.10);
        }
        .f-row{
          border-radius: 22px;
          padding: 14px 16px;
          background: rgba(255,255,255,0.56);
          backdrop-filter: blur(12px);
          display:flex;
          align-items:center;
          justify-content:space-between;
          gap: 16px;
        }
        .f-details{
          display:block;
        }
        .f-summary{
          list-style: none;
          cursor: pointer;
          user-select:none;
          outline: none;
        }
        .f-summary::-webkit-details-marker{ display:none; }
        .f-link{ text-decoration:none; display:block; color: inherit; }
        .f-left{ display:flex; align-items:flex-start; gap: 12px; min-width: 0; }
        .f-dot{
          width: 10px;
          height: 10px;
          border-radius: 999px;
          margin-top: 7px;
          background: linear-gradient(90deg, var(--dot1), var(--dot2));
          box-shadow: 0 10px 22px rgba(2,6,23,0.12);
          flex: 0 0 auto;
        }
        .f-title{
          font-weight: 950;
          font-size: 1.05rem;
          letter-spacing: -0.01em;
          line-height: 1.2;
          color: var(--text);
        }
        .f-sub{
          margin-top: 4px;
          color: var(--muted);
          font-size: 0.95rem;
          line-height: 1.35;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          max-width: 820px;
        }
        .f-actions{
          flex: 0 0 auto;
          display:flex;
          gap: 10px;
          align-items:center;
        }
        .f-badge{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          width: 240px;
          padding: 6px 10px;
          border-radius: 999px;
          border: 1px solid rgba(15,23,42,0.10);
          background: rgba(255,255,255,0.70);
          font-weight: 900;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          color: var(--text);
        }
        .f-cta{
          display:inline-flex;
          align-items:center;
          justify-content:center;
          padding: 8px 12px;
          border-radius: 999px;
          border: 1px solid transparent;
          background:
            linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.82)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.26), rgba(var(--accent2-rgb),0.26)) border-box;
          font-weight: 900;
          color: var(--text);
          white-space: nowrap;
        }
        .f-cta:after{
          content: "＋";
          margin-left: 6px;
          color: rgba(15,23,42,0.55);
          font-weight: 950;
        }
        .f-row-wrap[open] .f-cta:after{ content:"－"; }
        .f-expand{
          padding: 0 16px 16px 16px;
        }

        .pdf-card{
          border-radius: 22px;
          border: 1px solid transparent;
          padding: 10px;
          background:
            linear-gradient(rgba(255,255,255,0.72), rgba(255,255,255,0.66)) padding-box,
            linear-gradient(90deg, rgba(var(--accent1-rgb),0.22), rgba(var(--accent2-rgb),0.22)) border-box;
          box-shadow: 0 16px 48px rgba(2,6,23,0.08);
        }
        .pdf-card-inner{
          border-radius: 18px;
          overflow: hidden;
          background: rgba(255,255,255,0.88);
          border: 1px solid rgba(255,255,255,0.65);
          backdrop-filter: blur(12px);
        }
        .pdf-page{
          display:block;
          width:100%;
          height:auto;
          background: #fff;
        }
        .st-key-fullsol_prev button,
        .st-key-fullsol_next button{
          width: 52px;
          height: 52px;
          border-radius: 999px !important;
          border: 1px solid rgba(var(--accent1-rgb), 0.35) !important;
          background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(250,251,255,0.92)) !important;
          box-shadow: 0 8px 20px rgba(2,6,23,0.10);
          color: var(--text) !important;
          font-size: 1.35rem !important;
          font-weight: 900 !important;
          padding: 0 !important;
        }
        .st-key-fullsol_prev button:hover,
        .st-key-fullsol_next button:hover{
          transform: translateY(-1px);
          box-shadow: 0 12px 24px rgba(2,6,23,0.14);
        }
        .st-key-fullsol_prev button p,
        .st-key-fullsol_next button p{
          margin: 0 !important;
          line-height: 1 !important;
        }

        /* Mobile / small screens */
        @media (max-width: 860px){
          .block-container{
            padding-top: 0.9rem;
            padding-bottom: 2.0rem;
            padding-left: 0.95rem;
            padding-right: 0.95rem;
            max-width: 100%;
          }

          /* Stack Streamlit columns on narrow screens */
          div[data-testid="stHorizontalBlock"]{
            flex-direction: column !important;
            gap: 0.75rem !important;
          }
          div[data-testid="stHorizontalBlock"] > div[data-testid="column"]{
            width: 100% !important;
            flex: 1 1 100% !important;
          }

          /* Segmented controls: full width */
          [data-testid="stSegmentedControl"]{
            width: 100% !important;
          }
          [data-testid="stSegmentedControl"] button{
            flex: 1 1 0 !important;
          }

          .hero-title{ font-size: 1.55rem; }
          .hero-desc{ font-size: 0.95rem; }

          .spec-grid{ grid-template-columns: 1fr; }
          .tile-grid{ grid-template-columns: repeat(2, minmax(0, 1fr)); }

          .kv-grid{ grid-template-columns: 1fr; }
          .kv-k{ white-space: normal; }
          .kv-v{ overflow-wrap: anywhere; word-break: break-word; }

          .f-row{
            flex-direction: column;
            align-items: stretch;
            gap: 10px;
          }
          .f-actions{
            width: 100%;
            justify-content: space-between;
          }
          .f-badge{
            width: 100%;
            max-width: none;
          }
          .f-sub{
            white-space: normal;
            max-width: none;
          }
          .v-grid{ grid-template-columns: 1fr; }

          .ip-card{ min-height: 0; margin-bottom: 14px; }
          .ip-avatar{ width: 82px; height: 82px; padding: 6px; }
          .ip-latin{
            max-width: none;
            white-space: normal;
            overflow: visible;
            text-overflow: initial;
          }
          .ip-k{ white-space: normal; }
          .chip{ white-space: normal; }
          [data-testid="stButton"] > button{
            min-height: 44px;
            font-size: 1rem;
            touch-action: manipulation;
            -webkit-tap-highlight-color: transparent;
          }
          .st-key-fullsol_prev button,
          .st-key-fullsol_next button{
            width: 48px;
            height: 48px;
            min-height: 48px;
          }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <style>
        :root{{
          --accent1: {accent1};
          --accent2: {accent2};
          --accent3: {accent3};
          --tint: {tint};
          --accent1-rgb: {r1},{g1},{b1};
          --accent2-rgb: {r2},{g2},{b2};
          --accent3-rgb: {r3},{g3},{b3};
        }}
        .pill{{
          border-color: rgba({r1},{g1},{b1},0.22);
          background: var(--tint);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    logo_mask_src = ""
    if LOGO_SVG_PATH.exists():
        try:
            logo_cache_buster = LOGO_SVG_PATH.stat().st_mtime
        except Exception:
            logo_cache_buster = None
        logo_mask_src = load_image_data_uri(str(LOGO_SVG_PATH), logo_cache_buster)
    if not logo_mask_src and LOGO_ICON_PATH.exists():
        try:
            wm_cache_buster = LOGO_ICON_PATH.stat().st_mtime
        except Exception:
            wm_cache_buster = None
        logo_mask_src = load_image_data_uri(str(LOGO_ICON_PATH), wm_cache_buster)

    # Load hero visual: use clean hero first, then fallback assets
    hero_art_src = ""
    for candidate in [
        HERO_ART_PATH,
        resource_path("Final/示范.png"),
        resource_path("Final/logo 1.png"),
        resource_path("docs/assets/hero.png"),
    ]:
        if candidate.exists():
            try:
                _cb = candidate.stat().st_mtime
            except Exception:
                _cb = None
            hero_art_src = load_image_data_uri(str(candidate), _cb)
            if hero_art_src:
                break

    title = "人类健康与营养解决方案" if ui_lang == "CN" else "Human Health & Nutrition Solutions"
    desc_html = (
        "<div class='hero-desc'>"
        "<div><strong>Tailored for Global Brands</strong></div>"
        "<div>Built on scientific rigor, consumer insights, and global expertise.</div>"
        "<div style='margin-top:8px'>From proprietary probiotic strains to advanced formulations and end-to-end delivery.</div>"
        "</div>"
    ) if ui_lang == "EN" else (
        "<div class='hero-desc'>"
        "<div><strong>服务全球品牌</strong></div>"
        "<div>以科学为基础，以洞察为导向，以专业能力贯穿全流程</div>"
        "<div style='margin-top:8px'>从自有益生菌菌株到配方开发与商业化落地</div>"
        "</div>"
    )

    visual_html = "<div class='hero-visual'></div>"
    if hero_art_src:
        safe_art_src = html.escape(hero_art_src, quote=True)
        wm_html = ""
        if logo_mask_src:
            safe_wm_src = html.escape(logo_mask_src, quote=True)
            wm_html = f"<img class='hero-visual-wm' src='{safe_wm_src}' alt='' />"
        chips_html = (
            "<div class='hero-visual-meta'>"
            "<span class='hero-visual-chip'>Science-led</span>"
            "<span class='hero-visual-chip'>Global Delivery</span>"
            "</div>"
            if ui_lang == "EN"
            else
            "<div class='hero-visual-meta'>"
            "<span class='hero-visual-chip'>科学驱动</span>"
            "<span class='hero-visual-chip'>全球交付</span>"
            "</div>"
        )
        visual_html = (
            "<div class='hero-visual'>"
            f"<img class='hero-photo' src='{safe_art_src}' alt='' />"
            f"{wm_html}{chips_html}"
            "</div>"
        )

    if "wec_series" not in st.session_state:
        st.session_state["wec_series"] = _SERIES_OPTIONS[0]
    if st.session_state.get("wec_series") not in _SERIES_OPTIONS:
        st.session_state["wec_series"] = _SERIES_OPTIONS[0]

    with st.container(border=True):
        _W_PATH = "M241.763 12.967C231.238 15.523 204.444 32.744 189.131 53.148C173.817 73.557 164.889 91.103 138.42 109.275C138.42 109.275 125.677 118.88 116.428 118.867C107.161 118.848 102.295 107.564 105.529 100.772C105.529 100.772 107.486 93.98 111.626 86.945C115.765 79.929 125.741 66.651 121.551 63.888C121.551 63.888 120.956 60.49 108.003 75.483C95.05 90.472 64.116 121.73 35.415 136.069C33.645 136.952 31.98 137.679 30.416 138.26C23.125 140.963 18.011 140.51 14.549 138.26C6.709 133.169 7.317 118.871 10.222 111.201C14.7 99.395 32.552 67.817 71.782 35.928C71.782 35.928 81.534 29.396 78.09 23.573C74.645 17.737 60.8 23.528 60.8 23.528C60.8 23.528 32.552 38.16 17.563 54.736C11.836 61.075 6.668 62.786 3.443 60.878C-0.66 58.454 -1.611 50.18 3.443 38.146C4.815 34.88 6.631 31.335 8.946 27.553V27.566C8.946 27.566 25.146 8.013 58.751 1.866C66.101 0.516 74.28 -0.188 83.309 0.095C83.309 0.095 87.736 0.095 92.337 1.866C95.818 3.201 99.395 5.552 101.229 9.682C105.478 19.256 96.326 28.184 96.326 28.184C96.326 28.184 47.266 74.998 32.73 106.979C32.73 106.979 29.853 114.306 36.142 111.64C42.431 108.969 79.105 88.871 108.438 59.232C108.438 59.232 122.15 43.594 131.079 50.61C140.002 57.631 125.338 76.773 125.338 76.773C125.338 76.773 108.762 95.274 117.366 103.571C125.969 111.85 149.447 84.119 176.612 33.714C187.219 14.069 199.747 5.278 211.086 1.866C221.583 -1.295 231.06 0.155 237.057 1.866C241.205 3.05 243.689 4.363 243.689 4.363C243.689 4.363 252.292 10.428 241.763 12.967Z"
        _logo_html = (
            "<div class='hero-logo'>"
            f"<svg class='hero-logo-w-svg' viewBox='0 0 247 141' fill='none' xmlns='http://www.w3.org/2000/svg'>"
            f"<path d='{_W_PATH}' fill='#D10025'/>"
            "</svg>"
            "<div class='hero-logo-divider'></div>"
            "<div class='hero-logo-text'>"
            "<span class='hero-logo-wordmark'>WECARE</span>"
            "<span class='hero-logo-sub'>PROBIOTICS &nbsp;&middot;&nbsp; SCIENCE &nbsp;&middot;&nbsp; SOLUTIONS</span>"
            "</div>"
            "</div>"
        )

        st.markdown(
            f"<div class='hero-wrap'>"
            "<div class='hero-content'>"
            f"{_logo_html}"
            "<div class='hero-head'>"
            f"<div class='hero-title'>{html.escape(title)}</div>"
            f"{desc_html}"
            "</div>"
            "</div>"
            f"{visual_html}"
            "</div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='hero-seg-row'></div>", unsafe_allow_html=True)
        nav_left, nav_right = st.columns([8.2, 1.8])
        with nav_left:
            label = "Wec 系列" if ui_lang == "CN" else "Wec Series"
            st.segmented_control(
                label,
                _SERIES_OPTIONS,
                key="wec_series",
                label_visibility="collapsed",
                width="stretch",
            )
        with nav_right:
            st.segmented_control(
                "语言",
                ["EN", "CN"],
                key="ui_lang",
                label_visibility="collapsed",
                width="stretch",
            )


def _png_bytes_to_data_uri(png_bytes: bytes) -> str:
    if not png_bytes:
        return ""
    b64 = base64.b64encode(png_bytes).decode("ascii")
    return f"data:image/png;base64,{b64}"


def _bytes_to_data_uri(mime: str, raw: bytes) -> str:
    if not raw:
        return ""
    b64 = base64.b64encode(raw).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _guess_image_mime(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".png":
        return "image/png"
    if ext in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if ext == ".svg":
        return "image/svg+xml"
    return "application/octet-stream"


@st.cache_data
def load_image_data_uri(path: str, _cache_buster: float | None = None) -> str:
    p = Path(path)
    if not p.exists():
        return ""
    try:
        raw = p.read_bytes()
    except Exception:
        return ""
    return _bytes_to_data_uri(_guess_image_mime(p), raw)


def _get_query_param_first(key: str) -> str:
    """兼容 st.query_params 与 experimental API，返回第一个 query param 值。"""
    try:
        qp = st.query_params  # type: ignore[attr-defined]
        if key not in qp:
            return ""
        v = qp.get(key)
        if isinstance(v, list):
            return str(v[0]) if v else ""
        return str(v) if v is not None else ""
    except Exception:
        try:
            qp = st.experimental_get_query_params()
            v = qp.get(key, [])
            return str(v[0]) if v else ""
        except Exception:
            return ""


def _clear_query_param(key: str) -> None:
    """Best-effort removal of a single query param without breaking other params."""
    try:
        qp = st.query_params  # type: ignore[attr-defined]
        if key in qp:
            qp.pop(key, None)
        return
    except Exception:
        pass

    try:
        qp = st.experimental_get_query_params()
        qp.pop(key, None)
        st.experimental_set_query_params(**qp)
    except Exception:
        pass


def _is_mobile_client() -> bool:
    """Best-effort mobile detection for rendering/perf tuning."""
    q = _get_query_param_first("mobile").strip().lower()
    if q in {"1", "true", "yes"}:
        return True
    if q in {"0", "false", "no"}:
        return False

    ua = ""
    try:
        ctx = getattr(st, "context", None)
        headers = getattr(ctx, "headers", None) if ctx is not None else None
        if headers:
            ua = str(headers.get("user-agent", ""))
    except Exception:
        ua = ""

    if not ua:
        return False
    return bool(re.search(r"iphone|ipad|android|mobile", ua, flags=re.I))


def _stat_cache_buster(path: Path) -> int | None:
    """Return a high-resolution cache key for local files."""
    try:
        st = path.stat()
        # st_mtime_ns is available on Python 3.3+; fall back to seconds if needed.
        return int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1_000_000_000)))
    except Exception:
        return None


@st.dialog("Strain Details")
def _show_weclac_strain_dialog(
    *,
    ui_lang: str,
    code: str,
    title: str,
    latin_name: str,
    feature: str,
    clinical: str,
    patent: str,
    spec: str,
    is_core: bool,
    icon_src: str,
    directions: List[str],
) -> None:
    t = (lambda cn, en: en) if ui_lang == "EN" else (lambda cn, en: cn)

    star = " <span class='detail-star'>★</span>" if is_core else ""
    title_html = html.escape(title)
    latin_html = ""
    if ui_lang == "EN" and latin_name:
        title_html = _format_sci_name_html(latin_name)
    elif latin_name:
        latin_html = f"<div class='detail-sub'>{_format_sci_name_html(latin_name)}</div>"

    feature_display = (
        feature.replace("，", " · ").replace(",", " · ") if ui_lang == "CN" else feature
    )
    feature_html = (
        f"<div class='detail-sub'>{html.escape(feature_display)}</div>" if feature_display else ""
    )
    badges = (
        f"<span class='tile-badge tile-badge-strong'>{html.escape(code)}</span>"
        + (
            f"<span class='tile-badge'>{html.escape(t('核心菌', 'Core strain'))}</span>"
            if is_core
            else ""
        )
    )

    st.markdown(
        (
            "<div class='detail-wrap'>"
            "<div class='detail-card'>"
            "<div style='display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap'>"
            f"<div><div class='detail-title'>{title_html}{star}</div>"
            + latin_html
            + feature_html
            + "</div>"
            f"<div style='display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end'>{badges}</div>"
            "</div>"
            "<div class='detail-grid'>"
            f"<div class='detail-avatar'><img src='{icon_src}' alt='{html.escape(code)}' /></div>"
            "<div class='kv-table'>"
            "<div class='kv-grid'>"
            f"<div class='kv-k'>{html.escape(t('产品特点', 'Strain Highlights'))}</div>"
            f"<div class='kv-v'>{html.escape(feature)}</div>"
            f"<div class='kv-k'>{html.escape(t('临床数量', 'Clinical Studies'))}</div>"
            f"<div class='kv-v'>{html.escape(clinical)}</div>"
            f"<div class='kv-k'>{html.escape(t('专利数量', 'Patents'))}</div>"
            f"<div class='kv-v'>{html.escape(patent)}</div>"
            f"<div class='kv-k'>{html.escape(t('规格', 'Specification (CFU)'))}</div>"
            f"<div class='kv-v'>{html.escape(spec)}</div>"
            "</div>"
            "</div>"
            "</div>"
            "</div>"
            "</div>"
        ),
        unsafe_allow_html=True,
    )

    if st.button(t("关闭", "Close"), type="secondary"):
        st.session_state.pop("weclac_open", None)
        st.rerun()


def _weclac_placeholder_svg_data_uri(accent1: str, accent2: str) -> str:
    a1 = (accent1 or "#7C3AED").strip()
    a2 = (accent2 or "#FF2D55").strip()
    svg = f"""<svg xmlns='http://www.w3.org/2000/svg' width='96' height='96' viewBox='0 0 96 96'>
  <defs>
    <linearGradient id='g' x1='0' y1='0' x2='1' y2='1'>
      <stop stop-color='{html.escape(a1)}' stop-opacity='0.95'/>
      <stop offset='1' stop-color='{html.escape(a2)}' stop-opacity='0.95'/>
    </linearGradient>
  </defs>
  <circle cx='48' cy='48' r='34' fill='none' stroke='url(#g)' stroke-width='8'/>
  <circle cx='48' cy='48' r='34' fill='none' stroke='white' stroke-opacity='0.55' stroke-width='2'/>
</svg>"""
    return _bytes_to_data_uri("image/svg+xml", svg.encode("utf-8"))


def _render_pdf_page_card(png_bytes: bytes) -> None:
    src = _png_bytes_to_data_uri(png_bytes)
    if not src:
        return
    st.markdown(
        (
            "<div class='pdf-card'>"
            "<div class='pdf-card-inner'>"
            f"<img class='pdf-page' src='{src}' />"
            "</div>"
            "</div>"
        ),
        unsafe_allow_html=True,
    )


def _render_series_selector() -> None:
    if "wec_series" not in st.session_state:
        st.session_state["wec_series"] = "WecLac"
    if st.session_state.get("wec_series") not in _SERIES_OPTIONS:
        st.session_state["wec_series"] = "WecLac"

    with st.container(border=True):
        left, right = st.columns([1, 5])
        with left:
            ui_lang = str(st.session_state.get("ui_lang", "CN")).strip().upper() or "CN"
            st.markdown("&nbsp;", unsafe_allow_html=True)
        with right:
            st.segmented_control(
                "Wec Series",
                _SERIES_OPTIONS,
                key="wec_series",
                label_visibility="collapsed",
            )


def _extract_strain_code(name: str) -> Tuple[str, str]:
    """从菌株名称末尾提取类似 'BLa80' 的代号，并返回 (base_name, code)。"""
    text = (name or "").strip()
    if not text:
        return "", ""
    m = re.search(r"([A-Za-z]{1,6}\d{2,3})$", text)
    if not m:
        return text, ""
    code = m.group(1)
    base = text[: -len(code)].strip() or text
    return base, code


def _render_weclac_page() -> None:
    ui_lang = str(st.session_state.get("ui_lang", "CN")).strip().upper() or "CN"
    t = (lambda cn, en: en) if ui_lang == "EN" else (lambda cn, en: cn)

    pptx_path = os.getenv("DESIGN_WECLAC_PPTX", "").strip() or str(PPT_WECLAC_PATH)
    p = Path(pptx_path)
    if not p.exists():
        st.error(f"未找到 `WecLac.pptx`：`{pptx_path}`")
        return

    cache_buster = _stat_cache_buster(p)

    data = load_weclac_catalog(str(p), ui_lang, cache_buster)
    strains = list(data.get("strains", [])) if isinstance(data, dict) else []
    if not strains:
        st.warning("未能从 `WecLac.pptx` 提取到可展示的信息。")
        return

    # 只保留菌株行（去噪），保持 PPT 顺序（默认最多 16 个，便于 4×4 展示）
    catalog: List[Dict[str, object]] = []
    seen_codes: set[str] = set()
    for item in strains:
        name = str(item.get("name", "")).strip()
        base_name, code = _extract_strain_code(name)
        if not code or code in seen_codes:
            continue
        seen_codes.add(code)
        enriched = dict(item)
        enriched["base_name"] = base_name
        enriched["code"] = code
        catalog.append(enriched)
        if len(catalog) >= 16:
            break

    if not catalog:
        st.warning("未能从 `WecLac.pptx` 提取到 12 个菌株信息。")
        return

    # 提取“功能方向”（来自 BLa80 的段落）
    functions: List[Dict[str, str]] = []
    for item in catalog:
        f = item.get("functions", [])
        if isinstance(f, list) and f:
            functions = [x for x in f if isinstance(x, dict)]
            break

    if ui_lang == "EN" and functions:
        area_aliases = {
            "Emotional & Cognitive Health": "Mental Health",
            "Emotional and Cognitive Health": "Mental Health",
            "Oral Health": "Dental & Oral Health",
            "Immune Health": "Immunological Health",
            "Infant and Child Health": "Infant Health",
            "Infant & Child Health": "Infant Health",
        }

        def map_area(label: str) -> str:
            raw = (label or "").strip()
            if not raw:
                return ""
            for k, v in area_aliases.items():
                if raw.lower() == k.lower():
                    return v
            return raw

        for f in functions:
            f["direction"] = map_area(str(f.get("direction", "")))

    directions: List[str] = []
    for f in functions:
        d = str(f.get("direction", "")).strip()
        if d:
            directions.append(d)
    directions = list(dict.fromkeys(directions))

    icon_dir = Path(os.getenv("DESIGN_WECLAC_IMAGES_DIR", "").strip() or str(WECLAC_IMAGES_DIR))

    theme = _SERIES_THEME.get("WecLac", {})
    placeholder_src = _weclac_placeholder_svg_data_uri(theme.get("accent1", ""), theme.get("accent2", ""))

    def resolve_icon_src(code: str) -> str:
        if not code:
            return placeholder_src
        for ext in (".png", ".jpg", ".jpeg", ".svg"):
            candidate = icon_dir / f"{code}{ext}"
            if candidate.exists():
                cb = _stat_cache_buster(candidate)
                return load_image_data_uri(str(candidate), cb) or placeholder_src
        return placeholder_src

    code_to_item = {str(it.get("code", "")).strip(): it for it in catalog if str(it.get("code", "")).strip()}

    def open_weclac(code: str) -> None:
        st.session_state["weclac_open"] = (code or "").strip()

    # Clicking the IP image uses query params; immediately transfer to session_state and clear
    # to avoid the feeling of “opening another page”.
    qp_open_code = _get_query_param_first("open_weclac").strip() or _get_query_param_first("strain").strip()
    if qp_open_code and qp_open_code in code_to_item:
        last_qp = str(st.session_state.get("weclac_last_qp", "")).strip()
        if last_qp != qp_open_code:
            st.session_state["weclac_last_qp"] = qp_open_code
            open_weclac(qp_open_code)
            _clear_query_param("open_weclac")
            _clear_query_param("strain")
            st.rerun()

    # Prefer session state (no visible widgets). Keep query-param fallback for very old links.
    open_code = (
        str(st.session_state.pop("weclac_open", "")).strip()
        or _get_query_param_first("open_weclac").strip()
        or _get_query_param_first("strain").strip()
    )
    if open_code and open_code in code_to_item:
        item = code_to_item[open_code]
        code = open_code
        name = str(item.get("name", "")).strip()
        base_name = str(item.get("base_name", "")).strip()
        feature = str(item.get("feature", "")).strip()
        clinical = str(item.get("clinical", "")).strip()
        patent = str(item.get("patent", "")).strip()
        spec = str(item.get("spec", "")).strip()
        is_core = code in WECLAC_CORE_CODES
        latin_name = _STRAIN_SCI_NAMES.get(code, "")
        if not latin_name and re.search(r"[A-Za-z]", base_name) and " " in base_name:
            latin_name = base_name
        src = resolve_icon_src(code)
        title = base_name or name
        _show_weclac_strain_dialog(
            ui_lang=ui_lang,
            code=code,
            title=title,
            latin_name=latin_name,
            feature=feature,
            clinical=clinical,
            patent=patent,
            spec=spec,
            is_core=is_core,
            icon_src=src,
            directions=directions,
        )
        _clear_query_param("open_weclac")
        _clear_query_param("strain")

    def build_open_href(code: str) -> str:
        # Preserve existing query params (series/lang) while adding open_weclac.
        params: Dict[str, List[str]] = {}
        try:
            for k in st.query_params.keys():
                v = st.query_params.get_all(k)
                params[str(k)] = [str(x) for x in v if str(x)]
        except Exception:
            params = {}
        params["lang"] = [ui_lang]
        params["series"] = ["WecLac"]
        params["open_weclac"] = [code]
        return "?" + urlencode(params, doseq=True)

    # 菌株 IP 网格（信息默认折叠）
    # - 默认 16 个：4 列 × 4 行更整齐
    # - 其他数量：回退 3 列布局
    cols_n = 4 if len(catalog) >= 16 else 3
    for row_start in range(0, len(catalog), cols_n):
        cols = st.columns(cols_n, gap="large")
        for col, item in zip(cols, catalog[row_start : row_start + cols_n]):
            with col:
                code = str(item.get("code", "")).strip()
                name = str(item.get("name", "")).strip()
                base_name = str(item.get("base_name", "")).strip()
                feature = str(item.get("feature", "")).strip()
                clinical = str(item.get("clinical", "")).strip()
                patent = str(item.get("patent", "")).strip()
                spec = str(item.get("spec", "")).strip()
                latin_name = _STRAIN_SCI_NAMES.get(code, "")
                if not latin_name and ui_lang == "EN" and re.search(r"[A-Za-z]", base_name) and " " in base_name:
                    latin_name = base_name

                src = resolve_icon_src(code)
                title = "" if ui_lang == "EN" else (base_name or name)
                code_line = ""
                if latin_name:
                    code_line += f"<span class='ip-latin'>{_format_sci_name_html(latin_name)}</span>"
                if code:
                    code_line += f"<span class='code-pill'>{html.escape(code)}</span>"
                title_html = f"<div class='ip-name'>{html.escape(title)}</div>" if title else ""
                href = html.escape(build_open_href(code), quote=True)
                st.markdown(
                    (
                        "<div class='ip-card'>"
                        f"<div class='ip-avatar'><a class='weclac-open' href='{href}' target='_self' aria-label='Open {html.escape(code)}'><img src='{src}' alt='{html.escape(code)}' /></a></div>"
                        f"<div class='ip-code'>{code_line}</div>"
                        f"{title_html}"
                        "</div>"
                    ),
                    unsafe_allow_html=True,
                )

    if directions:
        with st.container(border=True):
            st.markdown(f"**{t('功能方向', 'Supported Application Areas')}**")
            chip_html = "".join(f"<span class='chip'>{html.escape(d)}</span>" for d in directions)
            st.markdown(
                f"<div style='display:flex;gap:10px;flex-wrap:wrap'>{chip_html}</div>",
                unsafe_allow_html=True,
            )

            has_desc = any(str(f.get("desc", "")).strip() for f in functions)
            if has_desc:
                rows = ""
                for f in functions:
                    d = str(f.get("direction", "")).strip()
                    desc = str(f.get("desc", "")).strip()
                    if not d:
                        continue
                    rows += (
                        f"<div class='ip-k'>{html.escape(d)}</div>"
                        f"<div class='ip-v'>{_italicize_microbe_tokens_html(desc)}</div>"
                    )
                st.markdown(
                    "<details class='ip-details'>"
                    f"<summary>{html.escape(t('展开方向详情', 'Show area details'))}</summary>"
                    f"<div class='ip-kv' style='grid-template-columns: 190px 1fr;'>{rows}</div>"
                    "</details>",
                    unsafe_allow_html=True,
                )


def _render_formula_variants_html(direction: str, ui_lang: str) -> str:
    variants = _WECPRO_FORMULA_VARIANTS.get(direction, [])
    if not variants:
        return ""

    t = (lambda cn, en: en) if ui_lang == "EN" else (lambda cn, en: cn)
    cards: List[str] = []

    for v in variants:
        product = str((v.get("product", {}) or {}).get(ui_lang, "") or "").strip()
        tag = str((v.get("tag", {}) or {}).get(ui_lang, "") or "").strip()
        benefit = str((v.get("benefit", {}) or {}).get(ui_lang, "") or "").strip()
        core_cn = str(v.get("core_cn", "") or "").strip()
        codes = [str(x).strip() for x in (v.get("codes", []) or []) if str(x).strip()]

        core_html = "—"
        if ui_lang == "EN" and codes:
            parts: List[str] = []
            for code in codes:
                sci = _STRAIN_SCI_NAMES.get(code, "")
                if sci:
                    parts.append(f"<div>{_format_sci_name_html(sci)} {html.escape(code)}</div>")
                else:
                    parts.append(f"<div>{html.escape(code)}</div>")
            core_html = "".join(parts) if parts else "—"
        elif ui_lang != "EN" and core_cn:
            core_html = html.escape(core_cn)

        tag_html = f"<span class='v-tag'>{html.escape(tag)}</span>" if tag else ""
        cards.append(
            "<div class='v-box'>"
            f"<div class='v-title'>{html.escape(product)}{tag_html}</div>"
            f"<div class='v-meta'>{html.escape(t('健康功效', 'Benefits'))}</div>"
            f"<div class='v-text'>{html.escape(benefit) if benefit else '—'}</div>"
            f"<div class='v-meta'>{html.escape(t('核心配方', 'Core Formula'))}</div>"
            f"<div class='v-text'>{core_html}</div>"
            "</div>"
        )

    return "<div class='v-grid'>" + "".join(cards) + "</div>"


def _render_wecpro_formula_page() -> None:
    ui_lang = str(st.session_state.get("ui_lang", "CN")).strip().upper() or "CN"
    t = (lambda cn, en: en) if ui_lang == "EN" else (lambda cn, en: cn)

    pptx_path = os.getenv("DESIGN_WECPRO_FORMULA_PPTX", "").strip() or str(PPT_WECPRO_FORMULA_PATH)
    p = Path(pptx_path)
    if not p.exists():
        st.error(f"未找到 `Formula.pptx`：`{pptx_path}`")
        return

    cache_buster = _stat_cache_buster(p)

    items = load_wecpro_formula_catalog(str(p), cache_buster)
    if not items:
        st.warning("未能从 `Formula.pptx` 提取到可展示的信息。")
        return

    order = [d for d in _FORMULA_SLIDE_TO_DIRECTION.values() if d]
    direction_to_item: Dict[str, Dict[str, object]] = {}
    for item in items:
        direction = str(item.get("direction", "")).strip()
        if direction:
            direction_to_item[direction] = item

    # 7 行（横向列表）：点击“介绍 +”同页展开，不跳转
    for direction in order[:7]:
        item = direction_to_item.get(direction, {})
        product = str(item.get("product", "")).strip()
        benefit = str(item.get("benefit", "")).strip()
        strains = [str(s).strip() for s in item.get("strains", []) if str(s).strip()]  # type: ignore[arg-type]

        direction_label = (
            _CATEGORY_LABELS_EN.get(_clean_ui_key(direction), _clean_ui_key(direction))
            if ui_lang == "EN"
            else _clean_ui_key(direction)
        )
        benefit_text = (
            _WECPRO_FORMULA_BENEFIT_EN.get(direction, benefit)
            if ui_lang == "EN"
            else benefit
        )

        theme = _CATEGORY_THEME.get(direction, {})
        a1 = theme.get("accent1", _SERIES_THEME["WecPro® Formula"]["accent1"])
        a2 = theme.get("accent2", _SERIES_THEME["WecPro® Formula"]["accent2"])

        row1 = _rgba(a1, 0.16)
        row2 = _rgba(a2, 0.16)

        variants = _WECPRO_FORMULA_VARIANTS.get(direction, [])
        if variants:
            product_badge = f"<span class='f-badge'>{html.escape('3个配方' if ui_lang=='CN' else '3 Formulas')}</span>"
            benefit_html = html.escape(
                "高端款 / 基础款 / 高活性益生菌酸奶款" if ui_lang == "CN" else "Premium / Base / Active Probiotic Yogurt"
            )
        else:
            product_badge = f"<span class='f-badge'>{html.escape(product)}</span>" if product else ""
            benefit_html = html.escape(benefit_text) if benefit_text else "—"

        strains_html = "—"
        if strains:
            if ui_lang == "EN":
                parts: List[str] = []
                for line in strains:
                    codes = _extract_strain_codes(line)
                    for code in codes:
                        sci = _STRAIN_SCI_NAMES.get(code)
                        if sci:
                            parts.append(f"<div>{_format_sci_name_html(sci)} {html.escape(code)}</div>")
                        else:
                            parts.append(f"<div>{html.escape(code)}</div>")
                if parts:
                    strains_html = "".join(parts)
            else:
                strains_text = "、".join([s for s in strains if s])
                strains_html = html.escape(strains_text) if strains_text else "—"

        expand_html = (
            _render_formula_variants_html(direction, ui_lang)
            if variants
            else (
                "<div class='kv-table'>"
                "<div class='kv-grid'>"
                f"<div class='kv-k'>{html.escape(t('健康功效', 'Benefits'))}</div>"
                f"<div class='kv-v'>{benefit_html}</div>"
                f"<div class='kv-k'>{html.escape(t('核心配方', 'Core Formula'))}</div>"
                f"<div class='kv-v'>{strains_html}</div>"
                "</div>"
                "</div>"
            )
        )

        block = (
            f"<details class='f-row-wrap f-details' style='--row1:{row1};--row2:{row2};--dot1:{a1};--dot2:{a2};'>"
            "<summary class='f-summary'>"
            "<div class='f-row'>"
            "<div class='f-left'>"
            "<div class='f-dot'></div>"
            "<div style='min-width:0'>"
            f"<div class='f-title'>{html.escape(direction_label)}</div>"
            "</div>"
            "</div>"
            "<div class='f-actions'>"
            f"{product_badge}<span class='f-cta'>{html.escape(t('介绍', 'Details'))}</span>"
            "</div>"
            "</div>"
            "</summary>"
            "<div class='f-expand'>"
            + expand_html
            + "</div>"
            "</details>"
        )
        st.markdown(block, unsafe_allow_html=True)


@st.cache_resource
def _start_packaged_autoshutdown(timeout_seconds: int = 20) -> None:
    """打包为 .app 时：无会话一段时间后自动退出，避免残留进程导致“未响应”."""
    if not getattr(sys, "frozen", False):
        return

    from streamlit.runtime.runtime import Runtime, RuntimeState

    def monitor() -> None:
        last_active = time.time()
        while True:
            try:
                rt = Runtime.instance()
                state = rt.state
            except Exception:
                time.sleep(1)
                continue

            if state == RuntimeState.ONE_OR_MORE_SESSIONS_CONNECTED:
                last_active = time.time()
            elif (
                state == RuntimeState.NO_SESSIONS_CONNECTED
                and time.time() - last_active > timeout_seconds
            ):
                rt.stop()
                time.sleep(2)
                os._exit(0)

            time.sleep(2)

    threading.Thread(target=monitor, daemon=True).start()


def _render_packaged_quit_button() -> None:
    if not getattr(sys, "frozen", False):
        return

    with st.sidebar:
        if st.button("退出应用"):
            try:
                from streamlit.runtime.runtime import Runtime

                if Runtime.exists():
                    Runtime.instance().stop()
            finally:
                os._exit(0)


def main() -> None:
    st.set_page_config(page_title="人类健康与营养解决方案", layout="wide")
    _start_packaged_autoshutdown()
    _render_packaged_quit_button()

    # Debug / emergency: force-clear Streamlit caches via URL
    # Example: https://...streamlit.app/?clear_cache=1
    if _get_query_param_first("clear_cache").strip() in {"1", "true", "yes"}:
        if not bool(st.session_state.get("_did_clear_cache")):
            st.session_state["_did_clear_cache"] = True
            try:
                st.cache_data.clear()
            except Exception:
                pass
            try:
                st.cache_resource.clear()
            except Exception:
                pass
        _clear_query_param("clear_cache")
        st.rerun()

    # UI 语言：EN / CN（可通过 ?lang=CN 直达中文）
    lang_from_url = _get_query_param_first("lang").strip().upper()
    if "ui_lang" not in st.session_state:
        st.session_state["ui_lang"] = "CN" if lang_from_url == "CN" else "EN"
    if str(st.session_state.get("ui_lang", "EN")).strip().upper() not in {"CN", "EN"}:
        st.session_state["ui_lang"] = "EN"
    ui_lang = str(st.session_state.get("ui_lang", "EN")).strip().upper() or "EN"
    is_mobile = _is_mobile_client()

    # Wec 系列入口（WecLac / WecPro® Formula / WecPro® Solution）
    series_from_url = _get_query_param_first("series").strip()
    if "wec_series" not in st.session_state:
        st.session_state["wec_series"] = (
            series_from_url if series_from_url in _SERIES_OPTIONS else "WecLac"
        )
    if st.session_state.get("wec_series") not in _SERIES_OPTIONS:
        st.session_state["wec_series"] = (
            series_from_url if series_from_url in _SERIES_OPTIONS else "WecLac"
        )
    series = str(st.session_state.get("wec_series", "WecLac"))

    if series == "WecLac":
        _render_header(series=series, badge=series)
        _render_weclac_page()
        return

    if series == "WecPro® Formula":
        _render_header(series=series, badge=series)
        _render_wecpro_formula_page()
        return

    excel_url = os.getenv("DESIGN_EXCEL_URL", "").strip()
    excel_path = EXCEL_PATH
    if excel_url:
        downloaded = fetch_remote_excel(excel_url)
        if downloaded:
            excel_path = Path(downloaded)

    if excel_path is None or not excel_path.exists():
        st.error(
            "未找到‘产品配方设计*.xlsx’数据文件。\n"
            "- 请将 Excel 放在 app.py 同目录（或 .app 同级目录），并确保文件名不是以 ~$ 开头\n"
            "- 或设置环境变量 DESIGN_EXCEL 指定完整路径\n"
            "- 或设置环境变量 DESIGN_EXCEL_URL 指向可下载的 Excel 链接（用于在线托管自动更新）"
        )
        st.stop()

    try:
        cache_buster = excel_path.stat().st_mtime
    except Exception:
        cache_buster = None

    overview = load_product_overview(str(excel_path), cache_buster)

    # 优先使用 Formula&Solution 的“功能方向 / 应用场景”作为筛选数据源（最新）
    formula_pptx = os.getenv("DESIGN_FORMULA_PPTX", "").strip() or str(PPT_FORMULA_PATH)
    formula_cache_buster = None
    try:
        if Path(formula_pptx).exists():
            formula_cache_buster = Path(formula_pptx).stat().st_mtime
    except Exception:
        formula_cache_buster = None

    try:
        formula_scenarios = (
            load_formula_scenarios(formula_pptx, formula_cache_buster)
            if Path(formula_pptx).exists()
            else {}
        )
    except Exception:
        formula_scenarios = {}

    # Load scenario-level metadata from the design sheet as a reliable fallback.
    try:
        _design_mapping, design_meta, _design_main_order, design_sub_order = load_solution_design(
            str(excel_path), cache_buster
        )
    except Exception:
        _design_mapping, design_meta, _design_main_order, design_sub_order = {}, {}, [], {}

    if formula_scenarios:
        ordered_main = [
            _FORMULA_SLIDE_TO_DIRECTION[s]
            for s in sorted(_FORMULA_SLIDE_TO_DIRECTION.keys())
            if _FORMULA_SLIDE_TO_DIRECTION[s] in formula_scenarios
        ]
        available_main = [m for m in ordered_main if formula_scenarios.get(m)]
        if not available_main:
            available_main = sorted([k for k, v in formula_scenarios.items() if v])
    else:
        available_main = [m for m in _design_main_order if m in _design_mapping]
        formula_scenarios = {k: design_sub_order.get(k, []) for k in available_main}

    if not available_main:
        st.error("未能读取到功能方向数据，请检查 Excel 或 Formula&Solution 文件。")
        return

    if "filter_cat" not in st.session_state:
        st.session_state["filter_cat"] = available_main[0]
    if st.session_state["filter_cat"] not in available_main:
        st.session_state["filter_cat"] = available_main[0]

    def reset_sub() -> None:
        current_cat = str(st.session_state.get("filter_cat", "")).strip()
        options = formula_scenarios.get(current_cat, [])
        st.session_state["filter_sub"] = options[0] if options else ""

    if "filter_sub" not in st.session_state:
        reset_sub()

    t = (lambda cn, en: en) if ui_lang == "EN" else (lambda cn, en: cn)

    # 解决方案 PPT：CN 用于映射/定位页码；EN 用于英文标题/核心功能解析
    solutions_pptx_cn = resolve_solutions_pptx_path("CN")
    solutions_pptx_en = resolve_solutions_pptx_path("EN")
    solutions_pdf_en = resolve_solutions_pdf_path("EN") if ui_lang == "EN" else None

    solutions_deck: Dict[str, Dict[str, object]] = {}
    alias_map: Dict[str, str] = {}
    ppt_cache_buster = None
    try:
        if solutions_pptx_cn and solutions_pptx_cn.exists():
            for p in (solutions_pptx_cn, Path(formula_pptx)):
                if p.exists():
                    try:
                        ppt_cache_buster = max(ppt_cache_buster or 0, p.stat().st_mtime)
                    except Exception:
                        pass
            solutions_deck = load_ppt_solution_deck(str(solutions_pptx_cn), ppt_cache_buster)
            alias_map = build_scenario_to_solution_title(
                formula_pptx,
                str(solutions_pptx_cn),
                (ppt_cache_buster, "map-20260222"),
            )
    except Exception:
        solutions_deck = {}
        alias_map = {}

    # EN deck may have extra intro/summary pages; build bridge from CN solution order
    # (01-43) to EN PPT start slides.
    en_slide_by_cn: Dict[int, int] = {}
    en_title_by_cn: Dict[int, str] = {}
    en_pdf_page_by_cn: Dict[int, int] = {}
    cn_ordered: List[Tuple[int, str]] = sorted(
        [
            (int(v.get("slide_no", 0)), str(k))
            for k, v in solutions_deck.items()
            if int(v.get("slide_no", 0)) > 0
        ],
        key=lambda x: x[0],
    )
    en_ppt_starts: List[Tuple[int, str]] = []
    if ui_lang == "EN" and solutions_pptx_en and solutions_pptx_en.exists():
        try:
            en_cache_buster = solutions_pptx_en.stat().st_mtime
        except Exception:
            en_cache_buster = None
        try:
            en_ppt_starts = load_ppt_solution_start_slides(
                str(solutions_pptx_en), (en_cache_buster, "starts-20260222")
            )
            for (cn_slide_no, _cn_title), (en_slide_no, en_title) in zip(cn_ordered, en_ppt_starts):
                en_slide_by_cn[cn_slide_no] = en_slide_no
                if en_title:
                    en_title_by_cn[cn_slide_no] = str(en_title).strip()
        except Exception:
            en_slide_by_cn = {}
            en_title_by_cn = {}
            en_ppt_starts = []

    scenario_bridge: Dict[Tuple[str, str], Dict[str, object]] = {}

    # Header (color matches the selected 功能方向)
    current_cat = _clean_ui_key(st.session_state.get("filter_cat", ""))
    badge_label = current_cat if ui_lang == "CN" else _CATEGORY_LABELS_EN.get(current_cat, current_cat)
    _render_header(series=series, category=current_cat, badge=badge_label)

    scenario_title_en: Dict[str, str] = {}
    scenario_label_en: Dict[str, str] = {}
    pdf_titles_en: Dict[int, str] = {}
    pdf_starts_en: List[Tuple[int, str]] = []
    if ui_lang == "EN" and solutions_pdf_en and solutions_pdf_en.exists():
        try:
            pdf_cache_key_en = solutions_pdf_en.stat().st_mtime
        except Exception:
            pdf_cache_key_en = None
        try:
            pdf_titles_en = load_pdf_solution_titles(str(solutions_pdf_en), pdf_cache_key_en)
            pdf_starts_en = load_pdf_solution_start_pages(str(solutions_pdf_en), pdf_cache_key_en)
        except Exception:
            pdf_titles_en = {}
            pdf_starts_en = []

    # Build CN->EN PDF page bridge separately; do not mix PDF page numbers into PPT slide mapping.
    if ui_lang == "EN" and pdf_starts_en and cn_ordered:
        en_pdf_page_by_cn = {}
        for (cn_slide_no, _cn_title), (en_pdf_page, _en_title) in zip(cn_ordered, pdf_starts_en):
            en_pdf_page_by_cn[cn_slide_no] = en_pdf_page

    # Canonical bridge keyed by (category_cn, scenario_cn), using Formula order as 01-43.
    ordered_cat_list = [d for _, d in sorted(_FORMULA_SLIDE_TO_DIRECTION.items())]
    ordered_formula_pairs: List[Tuple[str, str]] = []
    for cat_name in ordered_cat_list:
        for scen_name in formula_scenarios.get(cat_name, []) or []:
            scen_clean = str(scen_name).strip()
            if scen_clean:
                ordered_formula_pairs.append((cat_name, scen_clean))

    for idx, (cat_name, scen_name) in enumerate(ordered_formula_pairs, start=1):
        row: Dict[str, object] = {"seq": idx}

        # CN anchor:
        # - EN mode: keep strict sequence anchor to align all blocks with EN PDF order.
        # - CN mode: preserve semantic matching by scenario title.
        if ui_lang == "EN" and idx - 1 < len(cn_ordered):
            cn_slide_no, cn_title = cn_ordered[idx - 1]
            row["cn_slide_no"] = cn_slide_no
            row["cn_title"] = cn_title
        else:
            mk = alias_map.get(scen_name) or alias_map.get(_normalize_match_key(scen_name)) or ""
            if mk and mk in solutions_deck:
                cn_slide_no = int(solutions_deck[mk].get("slide_no", 0))
                if cn_slide_no:
                    row["cn_slide_no"] = cn_slide_no
                    row["cn_title"] = mk
            elif idx - 1 < len(cn_ordered):
                cn_slide_no, cn_title = cn_ordered[idx - 1]
                row["cn_slide_no"] = cn_slide_no
                row["cn_title"] = cn_title

        # EN PPT slide mapping for top textual blocks.
        if idx - 1 < len(en_ppt_starts):
            en_slide_no, en_title = en_ppt_starts[idx - 1]
            row["en_slide_no"] = en_slide_no
            row["en_title"] = str(en_title or "").strip()
        else:
            cn_slide_no = int(row.get("cn_slide_no", 0) or 0)
            if cn_slide_no:
                en_slide_no = en_slide_by_cn.get(cn_slide_no, 0)
                if en_slide_no:
                    row["en_slide_no"] = en_slide_no
                    row["en_title"] = str(en_title_by_cn.get(cn_slide_no, "") or "").strip()

        # EN PDF page mapping for Full Solution preview/download.
        if idx - 1 < len(pdf_starts_en):
            en_pdf_page, en_pdf_title = pdf_starts_en[idx - 1]
            row["en_pdf_page"] = en_pdf_page
            if not str(row.get("en_title", "") or "").strip() and en_pdf_title:
                row["en_title"] = str(en_pdf_title).strip()
        else:
            cn_slide_no = int(row.get("cn_slide_no", 0) or 0)
            if cn_slide_no:
                en_pdf_page = en_pdf_page_by_cn.get(cn_slide_no, 0)
                if en_pdf_page:
                    row["en_pdf_page"] = en_pdf_page

        scenario_bridge[(cat_name, scen_name)] = row

    def _format_cat(v: object) -> str:
        s = _clean_ui_key(v)
        return _CATEGORY_LABELS_EN.get(s, s) if ui_lang == "EN" else s

    with st.container(border=True):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.selectbox(
                t("功能方向", "Health Area"),
                available_main,
                key="filter_cat",
                on_change=reset_sub,
                format_func=_format_cat,
            )
        with col2:
            sub_options = formula_scenarios.get(st.session_state["filter_cat"], [])
            if not sub_options:
                st.selectbox(t("应用场景", "Supported Application Areas"), [""], key="filter_sub")
            else:
                if st.session_state.get("filter_sub") not in sub_options:
                    st.session_state["filter_sub"] = sub_options[0]

                if ui_lang == "EN":
                    current_cat_for_sub = _clean_ui_key(st.session_state.get("filter_cat", ""))
                    for scen in sub_options:
                        bridge = scenario_bridge.get((current_cat_for_sub, str(scen).strip()), {})
                        mk = alias_map.get(scen) or alias_map.get(_normalize_match_key(scen)) or scen
                        en_title = str(bridge.get("en_title", "") or "").strip()
                        cn_slide_no = int(bridge.get("cn_slide_no", 0) or 0)
                        if not en_title and cn_slide_no:
                            en_title = en_title_by_cn.get(cn_slide_no, "").strip()
                        if not en_title and cn_slide_no and pdf_titles_en:
                            en_pdf_page_no = en_pdf_page_by_cn.get(cn_slide_no, 0)
                            if en_pdf_page_no:
                                en_title = str(pdf_titles_en.get(en_pdf_page_no, "")).strip()
                        if not en_title:
                            en_title = mk
                        scenario_title_en[scen] = en_title
                        # EN 与 CN 一致：不展示 01-43 编号，仅展示场景标题
                        scenario_label_en[scen] = en_title

                def _format_sub(v: object) -> str:
                    s = str(v)
                    return scenario_label_en.get(s, s) if ui_lang == "EN" else s

                # Keep original order from Formula&Solution for both CN and EN.
                sub_options_sorted = sub_options
                if st.session_state.get("filter_sub") not in sub_options_sorted and sub_options_sorted:
                    st.session_state["filter_sub"] = sub_options_sorted[0]
                st.selectbox(
                    t("应用场景", "Supported Application Areas"),
                    sub_options_sorted,
                    key="filter_sub",
                    format_func=_format_sub,
                )

        cat = _clean_ui_key(st.session_state.get("filter_cat", ""))
        sub = str(st.session_state.get("filter_sub", "")).strip() or (sub_options[0] if sub_options else "")
        cat_label = _CATEGORY_LABELS_EN.get(cat, cat) if ui_lang == "EN" else cat
        sub_label = scenario_title_en.get(sub, sub) if ui_lang == "EN" else sub
        st.markdown(
            f'<div class="pill">{html.escape(cat_label)} · {html.escape(sub_label)}</div>',
            unsafe_allow_html=True,
        )

    selected_bridge = scenario_bridge.get((cat, sub), {})
    match_key = (
        str(selected_bridge.get("cn_title", "")).strip()
        or alias_map.get(sub)
        or alias_map.get(_normalize_match_key(sub))
        or sub
    )
    ppt_solution = solutions_deck.get(match_key)
    selected_cn_slide_no = int(selected_bridge.get("cn_slide_no", 0) or 0)
    selected_en_slide_no = int(selected_bridge.get("en_slide_no", 0) or 0)
    selected_en_pdf_page = int(selected_bridge.get("en_pdf_page", 0) or 0)
    selected_seq_no = int(selected_bridge.get("seq", 0) or 0)
    if not selected_cn_slide_no and ppt_solution:
        selected_cn_slide_no = int(ppt_solution.get("slide_no", 0))  # type: ignore[arg-type]
    if not selected_en_slide_no and selected_cn_slide_no:
        selected_en_slide_no = en_slide_by_cn.get(selected_cn_slide_no, selected_cn_slide_no)
    overview_block: Dict[str, object] = {}
    if ui_lang == "EN" and selected_en_slide_no and solutions_pptx_en and solutions_pptx_en.exists():
        try:
            try:
                en_cache_buster = solutions_pptx_en.stat().st_mtime
            except Exception:
                en_cache_buster = None
            en_lines = load_pptx_slide_lines(str(solutions_pptx_en), selected_en_slide_no, en_cache_buster)
            if en_lines:
                overview_block = _parse_ppt_overview(en_lines)
        except Exception:
            overview_block = {}
    elif ppt_solution:
        try:
            overview_lines = list(ppt_solution.get("overview_lines", []))  # type: ignore[arg-type]
            overview_block = _parse_ppt_overview(overview_lines)
        except Exception:
            overview_block = {}

    overview_info = overview.get(cat, {})
    overview_name = str(overview_info.get("name", "")).strip()
    overview_formula = str(overview_info.get("core_formula", "")).strip()
    overview_display_name = _ensure_wecpro_registered(overview_name)

    def _render_full_solution_section() -> bool:
        with st.container(border=True):
            st.markdown(f"### {t('完整解决方案', 'Full Solution')}")
            if ui_lang != "EN" and (not solutions_pptx_cn or not solutions_pptx_cn.exists()):
                st.info(
                    t(
                        "未找到解决方案 PPT：\n"
                        "- 请将 PPT 放到 `Design/Final/`\n"
                        "- 或设置环境变量 `DESIGN_SOLUTIONS_PPTX` 指向 PPT 路径",
                        "Solutions PPT not found:\n"
                        "- Put the PPTX into `Design/Final/`\n"
                        "- Or set env var `DESIGN_SOLUTIONS_PPTX` to a local path",
                    )
                )
            elif ui_lang != "EN" and not ppt_solution:
                st.warning(
                    t(
                        "该应用场景未匹配到 PPT 解决方案内容（请确认名称一致或更新映射）。",
                        "No matching solution found in the PPT deck (please verify names or update the mapping).",
                    )
                )
            else:
                pdf_path = resolve_solutions_pdf_path(ui_lang)
                if not pdf_path:
                    st.caption(
                        t(
                            "未找到解决方案 PDF（可将 PDF 放到 `Design/Final/`，或设置环境变量 `DESIGN_SOLUTIONS_PDF` 指向 PDF 路径）。",
                            "Solutions PDF not found (put it into `Design/Final/`, or set env var `DESIGN_SOLUTIONS_PDF` / `DESIGN_SOLUTIONS_PDF_EN`).",
                        )
                    )
                else:
                    try:
                        pdf_stat = pdf_path.stat()
                        pdf_cache_buster = pdf_stat.st_mtime
                    except Exception:
                        pdf_cache_buster = None

                    if ui_lang == "EN":
                        cn_slide_no = selected_cn_slide_no or (
                            int(ppt_solution.get("slide_no", 0)) if ppt_solution else 0  # type: ignore[arg-type]
                        )
                        render_slide_no = selected_en_pdf_page
                        if not render_slide_no and selected_seq_no and len(pdf_starts_en) >= selected_seq_no:
                            render_slide_no = int(pdf_starts_en[selected_seq_no - 1][0])
                        if not render_slide_no and cn_slide_no:
                            render_slide_no = en_pdf_page_by_cn.get(cn_slide_no, 0)
                        if not render_slide_no:
                            # Fallback for exceptional cases where EN PDF starts are unavailable.
                            render_slide_no = selected_en_slide_no or (
                                en_slide_by_cn.get(cn_slide_no, cn_slide_no) if cn_slide_no else 0
                            )
                        if not render_slide_no:
                            st.warning(
                                "No matching English solution page was found for this scenario in the PDF."
                            )
                            return False
                    else:
                        cn_slide_no = selected_cn_slide_no or int(ppt_solution.get("slide_no", 1))  # type: ignore[arg-type]
                        render_slide_no = cn_slide_no

                    page_state_key = "full_solution_page_side"
                    if str(st.session_state.get(page_state_key, "")).strip() not in {"left", "right"}:
                        st.session_state[page_state_key] = "left"
                    mode_state_key = "full_solution_view_mode"
                    if str(st.session_state.get(mode_state_key, "")).strip() not in {"single", "dual"}:
                        st.session_state[mode_state_key] = "single" if is_mobile else "dual"

                    render_scale = 1.4 if is_mobile else 2.0
                    tool1, tool2, tool3 = st.columns([4, 1, 2])
                    with tool1:
                        mode_options = [t("单页", "Single"), t("双页", "Dual")]
                        default_mode = mode_options[1] if st.session_state.get(mode_state_key) == "dual" else mode_options[0]
                        mode_selected = st.segmented_control(
                            t("查看模式", "View mode"),
                            mode_options,
                            default=default_mode,
                            key="full_solution_mode_seg",
                            label_visibility="collapsed",
                            width="content",
                        )
                        st.session_state[mode_state_key] = "dual" if mode_selected == mode_options[1] else "single"
                    with tool2:
                        with st.popover(t("显示设置", "View settings"), icon=":material/tune:"):
                            render_scale = st.slider(
                                t("清晰度", "Quality"),
                                min_value=1.0,
                                max_value=2.2 if is_mobile else 3.0,
                                value=1.4 if is_mobile else 2.0,
                                step=0.5,
                            )

                    page1 = max(1, render_slide_no)
                    page2 = max(1, render_slide_no + 1)

                    # 下载：始终提供当前 2 页的 PDF
                    solution_title = str(match_key or sub)
                    if ui_lang == "EN":
                        solution_title = (
                            scenario_title_en.get(sub, "").strip()
                            or str(overview_block.get("title", "")).strip()
                            or solution_title
                        )
                    safe_title = _safe_filename_component(solution_title)
                    solution_index = selected_seq_no or max(1, (cn_slide_no + 1) // 2)
                    solution_filename = f"{solution_index:02d}-{safe_title}.pdf"
                    solution_pdf_bytes = build_solution_pdf_bytes(
                        str(pdf_path),
                        page1,
                        page2,
                        pdf_cache_buster,
                    )
                    with tool3:
                        if solution_pdf_bytes:
                            st.download_button(
                                t("下载 2 页 PDF", "Download 2-page PDF"),
                                data=solution_pdf_bytes,
                                file_name=solution_filename,
                                mime="application/pdf",
                                type="primary",
                                use_container_width=True,
                            )
                        else:
                            st.caption(t("（PDF 生成失败：请确认已安装 `pypdf`）", "(PDF build failed: please ensure `pypdf` is installed.)"))

                    st.caption(
                        t(
                            "点击两侧小箭头翻页；单页/双页可在上方切换。",
                            "Tap side arrows to flip pages; switch Single/Dual mode above.",
                        )
                    )
                    nav_l, preview_col, nav_r = st.columns([1.05, 7.9, 1.05], gap="small")
                    with nav_l:
                        left_clicked = st.button(
                            t("‹", "‹"),
                            key="fullsol_prev",
                            use_container_width=True,
                        )
                    with nav_r:
                        right_clicked = st.button(
                            t("›", "›"),
                            key="fullsol_next",
                            use_container_width=True,
                        )

                    current_side = str(st.session_state.get(page_state_key, "left"))
                    if left_clicked:
                        current_side = "left"
                    elif right_clicked:
                        current_side = "right"
                    st.session_state[page_state_key] = current_side

                    mode = str(st.session_state.get(mode_state_key, "single" if is_mobile else "dual"))
                    pages_to_render: Tuple[int, ...]
                    if mode == "dual":
                        pages_to_render = (page1, page2)
                    elif current_side == "right":
                        pages_to_render = (page2,)
                    else:
                        pages_to_render = (page1,)

                    with st.spinner(t("正在加载页面...", "Loading pages...")):
                        page_images = render_pdf_pages_png(
                            str(pdf_path),
                            pages_to_render,
                            render_scale,
                            pdf_cache_buster,
                        )

                    if not page_images:
                        st.warning(
                            t(
                                "页面渲染失败：\n"
                                "- 请确认已安装依赖 `pymupdf`（重新运行 `run_app.command` 会自动安装）\n"
                                "- 或检查 PDF 文件是否完整/可打开",
                                "Render failed:\n"
                                "- Ensure `pymupdf` is installed (re-run `run_app.command` to auto-install)\n"
                                "- Or verify the PDF file is valid and can be opened",
                            )
                        )
                    else:
                        with preview_col:
                            if mode == "dual":
                                c1, c2 = st.columns(2, gap="large")
                                with c1:
                                    _render_pdf_page_card(page_images[0])
                                with c2:
                                    _render_pdf_page_card(page_images[1] if len(page_images) > 1 else page_images[0])
                            else:
                                _render_pdf_page_card(page_images[0])
        return True

    # 顺序调整：Full Solution 前移到“临床研究”和“规格”之前
    if not _render_full_solution_section():
        return

    with st.container(border=True):
        st.subheader(t("核心配方", "Core Formula"))
        display_name = overview_display_name
        formula_html = _colorize_solution_formula_html(overview_formula, ui_lang)
        if overview_name and overview_formula:
            sep = "：" if ui_lang == "CN" else ":"
            st.markdown(
                "<div class='core-formula-line'>"
                f"<span class='core-formula-name'>{html.escape(display_name)}</span>"
                f"<span class='core-formula-sep'>{html.escape(sep)} </span>"
                f"{formula_html}"
                "</div>",
                unsafe_allow_html=True,
            )
        elif overview_formula:
            st.markdown(f"<div class='core-formula-line'>{formula_html}</div>", unsafe_allow_html=True)
        elif overview_name:
            st.markdown(f"**{display_name}**")
        else:
            st.caption(t("（该功能方向暂无‘Sheet2’信息记录）", "(No record found for this health area.)"))

        highlights = [str(x).strip() for x in overview_block.get("highlights", []) if str(x).strip()]  # type: ignore[arg-type]
        if highlights:
            st.markdown(
                "<div class='core-func-title'>"
                "<span class='core-func-dot'></span>"
                f"<span>{html.escape(t('核心功能', 'Core Functions'))}</span>"
                "</div>",
                unsafe_allow_html=True,
            )
            items_html = "".join(
                f"<li>{_italicize_microbe_tokens_html(x)}</li>" for x in highlights[:4]
            )
            st.markdown(f"<ul class='core-func-list'>{items_html}</ul>", unsafe_allow_html=True)

    trial_lines = [str(x).strip() for x in overview_block.get("trials", []) if str(x).strip()]  # type: ignore[arg-type]
    trial_entries = _parse_trial_entries(trial_lines)
    if not trial_entries and isinstance(ppt_solution, dict):
        fallback_trial_lines: List[str] = []
        for key in ("overview_lines", "evidence_lines"):
            for raw in list(ppt_solution.get(key, [])):  # type: ignore[arg-type]
                line = str(raw).strip()
                if line and _is_ppt_trial_line(line):
                    fallback_trial_lines.append(line)
        if fallback_trial_lines:
            trial_entries = _parse_trial_entries(fallback_trial_lines)

    if not trial_entries:
        regs_text = str(
            ((design_meta.get(cat, {}) or {}).get(sub, {}) or {}).get("clinical_regs", "")
        ).strip()
        if regs_text:
            trial_entries = _parse_clinical_regs_entries(regs_text)

    if trial_entries:
        clinical_data_path = resolve_clinical_data_path()
        article_links: Dict[str, str] = {}
        if clinical_data_path:
            try:
                clinical_cache_buster = clinical_data_path.stat().st_mtime
            except Exception:
                clinical_cache_buster = None
            try:
                article_links = load_clinical_article_links(
                    str(clinical_data_path), clinical_cache_buster
                )
            except Exception:
                article_links = {}

        with st.container(border=True):
            st.subheader(t("临床研究", "Clinical Studies"))
            rows_html = ""
            for key, ids in trial_entries:
                badge_parts: List[str] = []
                for reg_id in ids:
                    rid = (reg_id or "").strip().replace(" ", "")
                    url = article_links.get(rid, "")
                    if url:
                        safe_url = html.escape(url, quote=True)
                        badge_parts.append(
                            "<a class='tile-badge tile-badge-link' "
                            f"href='{safe_url}' target='_blank' rel='noopener noreferrer'>"
                            f"{html.escape(reg_id)}</a>"
                        )
                    else:
                        badge_parts.append(f"<span class='tile-badge'>{html.escape(reg_id)}</span>")
                badges = "".join(badge_parts)
                rows_html += (
                    f"<div class='kv-k'>{html.escape(key)}</div>"
                    "<div class='kv-v'>"
                    f"<div style='display:flex;gap:8px;flex-wrap:wrap'>{badges}</div>"
                    "</div>"
                )
            st.markdown(
                "<div class='kv-table'>"
                "<div class='kv-grid clinical-grid'>"
                f"{rows_html}"
                "</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    capsule_path = resolve_capsule_details_path()
    capsule_specs: List[Dict[str, str]] = []
    if capsule_path:
        try:
            capsule_cache_buster = capsule_path.stat().st_mtime
        except Exception:
            capsule_cache_buster = None

        try:
            capsule_details = load_capsule_details(str(capsule_path), ui_lang, capsule_cache_buster)
        except Exception:
            capsule_details = {}

        cap_candidates = list(capsule_details.get(cat, {}).keys())
        cap_query = sub
        if ui_lang == "EN":
            cap_query = str(selected_bridge.get("cn_title", "")).strip() or cap_query
        cap_key = _pick_capsule_scenario(cap_query, cap_candidates)
        cap_record = capsule_details.get(cat, {}).get(cap_key) if cap_key else None
        capsule_specs = list(cap_record.get("specs", [])) if isinstance(cap_record, dict) else []

    if capsule_specs:
        with st.container(border=True):
            st.subheader(t("规格", "Specifications"))
            clinical_label = t("临床菌配方", "Clinical Strain Formula")
            excipient_label = t("功能性辅料", "Functional Excipients")
            dosage_form_label = t("剂型", "Dosage form")
            capsule_label = t("胶囊", "Capsule")
            granule_label = t("颗粒剂 / 粉剂", "Granules / Powder Formulation")
            sep = "：" if ui_lang == "CN" else ":"

            clinical_bases: List[str] = []
            for spec in capsule_specs:
                base, dose = _parse_capsule_clinical(spec.get("clinical", ""))
                candidate = base
                # 某些来源会是“Clinical Strain Formula: LRa05+...”，此时取冒号后正文。
                if dose and re.search(r"(clinical|临床菌|formula|blend|配方)", base, flags=re.IGNORECASE):
                    candidate = dose
                if candidate:
                    clinical_bases.append(candidate)

            base_unique = [b for b in dict.fromkeys(clinical_bases) if b]
            if len(base_unique) == 1:
                clinical_value = base_unique[0]
            else:
                clinical_value = " / ".join(base_unique[:2]) if base_unique else ""
            clinical_product_name = overview_display_name
            if not clinical_product_name and overview_formula:
                clinical_product_name = _ensure_wecpro_registered(
                    re.split(r"[:：]", overview_formula, maxsplit=1)[0].strip()
                )
            product_html = (
                _format_tm_sup_html(clinical_product_name, add_if_missing=True)
                if clinical_product_name
                else ""
            )

            def _codes_with_plus_html(codes: List[str]) -> str:
                if not codes:
                    return ""
                out: List[str] = []
                for idx, code in enumerate(codes):
                    if idx > 0:
                        out.append("<span class='spec-plus'>+</span>")
                    out.append(f"<span class='spec-code'>{html.escape(code)}</span>")
                return "".join(out)

            clinical_codes = _extract_strain_codes(clinical_value)
            core_codes = _extract_strain_codes(overview_formula)
            extra_codes = [c for c in clinical_codes if c not in set(core_codes)]
            extra_codes_html = _codes_with_plus_html(extra_codes)
            all_codes_html = _codes_with_plus_html(clinical_codes)
            if product_html:
                product_badge_html = f"<span class='spec-product-name'>{product_html}</span>"
                if extra_codes:
                    clinical_display_html = (
                        f"{product_badge_html}<span class='spec-plus'>+</span>{extra_codes_html}"
                    )
                elif clinical_codes:
                    clinical_display_html = (
                        f"{product_badge_html}<span class='spec-plus'>+</span>{all_codes_html}"
                    )
                else:
                    clinical_display_html = product_badge_html
            else:
                clinical_display_html = html.escape(clinical_value)

            def _extract_functional_exc_names(raw_text: str) -> List[str]:
                exc_items_raw = _split_capsule_excipients(str(raw_text or "").strip())
                out: List[str] = []
                seen: set[str] = set()
                for raw_item in exc_items_raw:
                    # Split mixed segments like "FOS / Inulin + Cranberry Powder"
                    parts = [
                        p.strip()
                        for p in re.split(r"(?:/|／|\+|＋|&|\band\b)", raw_item, flags=re.IGNORECASE)
                        if p and p.strip()
                    ]
                    if not parts:
                        parts = [raw_item]
                    for part in parts:
                        formatted = _strip_mass_units(_format_capsule_excipient_item(part, ui_lang))
                        name_only = _excipient_name_only(formatted)
                        if not name_only or _is_filler_excipient(name_only, ui_lang):
                            continue
                        key = name_only.lower()
                        if key in seen:
                            continue
                        seen.add(key)
                        out.append(name_only)
                return out

            # 功能性辅料以 120B 为基准展示；若缺失，则回退到首个可用规格。
            exc_120_names: List[str] = []
            for _spec in capsule_specs:
                _label = str(_spec.get("spec", ""))
                if re.search(r"(?i)\b120\s*B\b", _label):
                    exc_120_names = _extract_functional_exc_names(str(_spec.get("excipients", "")))
                    break

            if not exc_120_names:
                for _spec in capsule_specs:
                    exc_120_names = _extract_functional_exc_names(str(_spec.get("excipients", "")))
                    if exc_120_names:
                        break

            exc_text = (
                ("、".join(exc_120_names) if ui_lang == "CN" else ", ".join(exc_120_names))
                if exc_120_names
                else "—"
            )

            dosage_forms = t("胶囊 / 颗粒剂 / 粉剂", "Capsules / Granules / Powder Formulation")
            def _check_item_html(dose: str) -> str:
                return (
                    "<span class='spec-check-item'>"
                    "<span class='spec-check-dot'>✓</span>"
                    f"<span>{html.escape(dose)}</span>"
                    "</span>"
                )

            lines: List[str] = []
            if clinical_display_html:
                lines.append(
                    "<div class='spec-line'>"
                    f"<span class='spec-k'>{html.escape(clinical_label)}{html.escape(sep)}</span>"
                    f"<span class='spec-v-formula'>{clinical_display_html}</span>"
                    "</div>"
                )

            lines.append(
                "<div class='spec-line'>"
                f"<span class='spec-k'>{html.escape(excipient_label)}{html.escape(sep)}</span>"
                f"<span class='spec-v'>{html.escape(exc_text)}</span>"
                "</div>"
            )
            lines.append(
                "<div class='spec-line'>"
                f"<span class='spec-k'>{html.escape(dosage_form_label)}{html.escape(sep)}</span>"
                f"<span class='spec-v'>{html.escape(dosage_forms)}</span>"
                "</div>"
            )

            capsule_checks = "".join(
                [_check_item_html("120B"), _check_item_html("240B"), _check_item_html("480B")]
            )
            granule_checks = "".join(
                [_check_item_html("120B"), _check_item_html("300B"), _check_item_html("1000B")]
            )
            lines.append(
                "<div class='spec-line'>"
                f"<span class='spec-k'>{html.escape(capsule_label)}{html.escape(sep)}</span>"
                f"<span class='spec-checklist'>{capsule_checks}</span>"
                "</div>"
            )
            lines.append(
                "<div class='spec-line'>"
                f"<span class='spec-k'>{html.escape(granule_label)}{html.escape(sep)}</span>"
                f"<span class='spec-checklist'>{granule_checks}</span>"
                "</div>"
            )

            st.markdown("<div class='spec-list'>" + "".join(lines) + "</div>", unsafe_allow_html=True)

            st.caption(
                t(
                    "实际配方组合可根据客户需求进行定制化设计。",
                    "The actual formulation can be customized according to customer requirements.",
                )
            )

    # 客户展示版：不展示“配方设计池 / 说明书 / 临床注册号”等内部信息


if __name__ == "__main__":
    # 在打包为可执行文件后，通过 streamlit run 启动
    if getattr(sys, "frozen", False):
        import socket
        import streamlit.web.cli as stcli

        # 让 runner 脚本可以 import app 并调用 main()
        sys.modules.setdefault("app", sys.modules[__name__])

        runner_code = (
            "from app import main, _show_fatal_dialog, _write_fatal_log\n"
            "try:\n"
            "    main()\n"
            "except Exception as e:\n"
            "    log_path = _write_fatal_log(e)\n"
            "    extra = f\"\\n\\n日志：{log_path}\" if log_path else \"\"\n"
            "    _show_fatal_dialog(\n"
            "        \"WECARE 产品解决方案 启动失败\",\n"
            "        \"应用启动时发生错误。\\n\\n\"\n"
            "        \"常见原因：\\n\"\n"
            "        \"1) 机型架构不匹配（Intel 与 Apple 芯片）。\\n\"\n"
            "        \"2) macOS 版本过低。\\n\"\n"
            "        \"3) 文件被系统隔离（quarantine）。\\n\\n\"\n"
            "        \"可尝试：右键应用→打开；或在‘隐私与安全性’中点‘仍要打开’。\"\n"
            "        + extra,\n"
            "    )\n"
            "    raise\n"
        )
        with tempfile.NamedTemporaryFile(
            "w", suffix=".py", delete=False, encoding="utf-8"
        ) as tf:
            tf.write(runner_code)
            temp_path = tf.name

        def pick_port() -> int:
            preferred = int(os.environ.get("STREAMLIT_PORT", "8501"))

            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                try:
                    s.bind(("127.0.0.1", preferred))
                    return preferred
                except OSError:
                    pass

            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", 0))
                return int(s.getsockname()[1])

        port = pick_port()
        sys.argv = [
            "streamlit",
            "run",
            temp_path,
            "--global.developmentMode=false",
            "--server.headless=false",
            "--server.address=localhost",
            f"--server.port={port}",
        ]
        stcli.main()
    else:
        main()
