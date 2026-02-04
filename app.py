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
PPT_SOLUTIONS_PATH = resource_path("Final/43 Solutions解决方案中文版20260130.pptx")
PPT_SOLUTIONS_EN_PATH = resource_path("Final/43 Solutions解决方案英文版20260130.pptx")
PPT_FORMULA_PATH = resource_path("Final/Formula&Solution.pptx")
PDF_SOLUTIONS_PATH = resource_path("Final/43 Solutions解决方案中文版20260130.pdf")
PDF_SOLUTIONS_EN_PATH = resource_path("Final/43 Solutions解决方案英文版20260130.pdf")
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
    try:
        idx = lines.index("Solution")
    except ValueError:
        return ""
    if idx + 1 >= len(lines):
        return ""
    return lines[idx + 1].strip()


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

    alias: Dict[str, str] = {}
    cursor = 0
    for slide_no in sorted(_FORMULA_SLIDE_TO_DIRECTION.keys()):
        direction = _FORMULA_SLIDE_TO_DIRECTION[slide_no]
        scenarios = direction_to_scenarios.get(direction, [])
        if not scenarios:
            continue

        chunk = solution_titles_ordered[cursor : cursor + len(scenarios)]
        cursor += len(scenarios)
        remaining = list(chunk)

        for scen in scenarios:
            picked = _best_title_match(scen, remaining)
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
    "杆菌",
    "双歧杆菌",
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

    for line in content_lines:
        if not line or line in meta_lines or line == "Solution" or line == title:
            continue

        if re.match(r"^(核心辅料|其他辅料|辅料)\s*[:：]", line):
            excipients.append(line)
            continue
        if re.match(r"^(Key Excipients|Other Excipients|Excipients)\s*[:：]", line, flags=re.I):
            excipients.append(line)
            continue

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

        if any(hint in line for hint in _PPT_STRAIN_HINTS):
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
            "potato starch",
            "starch",
            "silicon dioxide",
            "magnesium stearate",
        }
        return key in fillers
    fillers_cn = ("阿拉伯胶", "马铃薯淀粉", "二氧化硅", "硬脂酸镁", "淀粉")
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
    pages: Tuple[int, int],
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
            "core_cn": "乳酸片球菌PA53、植物乳植杆菌Lp18、动物双歧杆菌乳亚种BLa36、凝结魏茨曼氏菌BC99",
            "codes": ["PA53", "Lp18", "BLa36", "BC99"],
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

        /* Segmented control (pills) */
        [data-testid="stSegmentedControl"]{
          background: rgba(255,255,255,0.72);
          border: 1px solid rgba(15,23,42,0.12);
          border-radius: 999px;
          padding: 4px;
          backdrop-filter: none;
        }
        [data-testid="stSegmentedControl"] button{
          border-radius: 999px !important;
          padding: 7px 12px !important;
          min-height: 36px !important;
          font-weight: 750 !important;
          border: 0 !important;
          background: transparent !important;
          color: var(--muted) !important;
          box-shadow: none !important;
          transition: all .15s ease;
        }
        [data-testid="stSegmentedControl"] button[aria-selected="true"],
        [data-testid="stSegmentedControl"] button[aria-pressed="true"]{
          background: rgba(var(--accent1-rgb),0.14) !important;
          border: 1px solid rgba(var(--accent1-rgb),0.22) !important;
          color: var(--text) !important;
          box-shadow: 0 8px 18px rgba(2,6,23,0.10) !important;
        }

        /* Download button */
        [data-testid="stDownloadButton"] button{
          background: var(--accent1);
          border: 1px solid rgba(15,23,42,0.10);
        }
        [data-testid="stDownloadButton"] button p{ color: #fff; font-weight: 600; }

        .hero-title{
          font-size: 1.85rem;
          font-weight: 800;
          letter-spacing: -0.02em;
          line-height: 1.15;
          margin: 0;
        }
        .hero-head{
          display:flex;
          align-items:flex-start;
          gap: 14px;
          position: relative;
          overflow: hidden;
        }
        .hero-head > *{
          position: relative;
          z-index: 1;
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
          color: var(--text);
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
          min-height: 280px;
          margin: 0 0 16px 0;
          display:flex;
          flex-direction:column;
        }
        .ip-link{ text-decoration:none; display:block; }
        .ip-avatar{
          width: 78px;
          height: 78px;
          border-radius: 22px;
          overflow: hidden;
          margin: 0 auto;
          display:flex;
          align-items:center;
          justify-content:center;
          box-sizing: border-box;
          padding: 8px;
          background:
            linear-gradient(135deg, rgba(255,255,255,0.78), rgba(255,255,255,0.35));
          border: 1px solid rgba(255,255,255,0.65);
          box-shadow: 0 10px 26px rgba(2,6,23,0.10);
        }
        .ip-avatar img{
          width: 100%;
          height: 100%;
          object-fit: contain;
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
          .ip-avatar{ width: 70px; height: 70px; padding: 7px; }
          .ip-latin{
            max-width: none;
            white-space: normal;
            overflow: visible;
            text-overflow: initial;
          }
          .ip-k{ white-space: normal; }
          .chip{ white-space: normal; }
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

    if logo_mask_src:
        st.markdown(
            f"""
            <style>
            /* Header watermark 'W' (right-aligned, clipped to header card) */
            @supports selector(:has(*)){{
            [data-testid="stVerticalBlockBorderWrapper"]:has(#wecare-hero-marker){{
              overflow: hidden;
            }}
            [data-testid="stVerticalBlockBorderWrapper"]:has(#wecare-hero-marker)::before{{
              content: "";
              position: absolute;
              inset: 0;
              pointer-events: none;
              opacity: 0.22;
              background: linear-gradient(
                135deg,
                rgba(15, 23, 42, 0.95),
                rgba(100, 116, 139, 0.95)
              );
              -webkit-mask-image: url("{logo_mask_src}");
              -webkit-mask-repeat: no-repeat;
              -webkit-mask-position: right 12px center;
              -webkit-mask-size: 520px auto;
              mask-image: url("{logo_mask_src}");
              mask-repeat: no-repeat;
              mask-position: right 12px center;
              mask-size: 520px auto;
              filter: blur(0.2px);
            }}
            @media (max-width: 720px){{
              [data-testid="stVerticalBlockBorderWrapper"]:has(#wecare-hero-marker)::before{{
                -webkit-mask-size: 360px auto;
                mask-size: 360px auto;
                -webkit-mask-position: right 8px center;
                mask-position: right 8px center;
                opacity: 0.18;
              }}
            }}
            }}

            /* Fallback for browsers without :has(): target the first bordered container (header). */
            @supports not selector(:has(*)){{
            [data-testid="stVerticalBlockBorderWrapper"]:first-of-type{{
              overflow: hidden;
            }}
            [data-testid="stVerticalBlockBorderWrapper"]:first-of-type::before{{
              content: "";
              position: absolute;
              inset: 0;
              pointer-events: none;
              opacity: 0.22;
              background: linear-gradient(
                135deg,
                rgba(15, 23, 42, 0.95),
                rgba(100, 116, 139, 0.95)
              );
              -webkit-mask-image: url("{logo_mask_src}");
              -webkit-mask-repeat: no-repeat;
              -webkit-mask-position: right 12px center;
              -webkit-mask-size: 520px auto;
              mask-image: url("{logo_mask_src}");
              mask-repeat: no-repeat;
              mask-position: right 12px center;
              mask-size: 520px auto;
              filter: blur(0.2px);
            }}
            @media (max-width: 720px){{
              [data-testid="stVerticalBlockBorderWrapper"]:first-of-type::before{{
                -webkit-mask-size: 360px auto;
                mask-size: 360px auto;
                -webkit-mask-position: right 8px center;
                mask-position: right 8px center;
                opacity: 0.18;
              }}
            }}
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

    with st.container(border=True):
        if logo_mask_src:
            st.markdown(
                "<span id='wecare-hero-marker' style='display:none' aria-hidden='true'></span>",
                unsafe_allow_html=True,
            )
        cols = st.columns([9, 2])
        with cols[0]:
            title = "WECARE 健康与营养解决方案" if ui_lang == "CN" else "WECARE Health & Wellness Solutions"
            desc_html = ""
            if ui_lang == "EN":
                desc_html = (
                    "<div class='hero-desc'>"
                    "<div><strong>Tailored for Global Brands</strong></div>"
                    "<div>Built on scientific rigor, consumer insights, and global expertise.</div>"
                    "<div style='margin-top:10px'>From proprietary probiotic strains to advanced formulations and end-to-end delivery.</div>"
                    "</div>"
                )
            else:
                desc_html = (
                    "<div class='hero-desc'>"
                    "<div><strong>服务全球品牌</strong></div>"
                    "<div>以科学为基础，以洞察为导向，以专业能力贯穿全流程</div>"
                    "<div style='margin-top:10px'>从自有益生菌菌株到配方开发与商业化落地</div>"
                    "</div>"
                )
            st.markdown(
                "<div class='hero-head'><div>"
                f"<div class='hero-title'>{html.escape(title)}</div>{desc_html}"
                "</div></div>",
                unsafe_allow_html=True,
            )

            if "wec_series" not in st.session_state:
                st.session_state["wec_series"] = _SERIES_OPTIONS[0]
            if st.session_state.get("wec_series") not in _SERIES_OPTIONS:
                st.session_state["wec_series"] = _SERIES_OPTIONS[0]

            label = "Wec 系列" if ui_lang == "CN" else "Wec Series"
            st.segmented_control(
                label,
                _SERIES_OPTIONS,
                key="wec_series",
                label_visibility="collapsed",
                width="content",
            )
        with cols[1]:
            st.segmented_control(
                "语言",
                ["CN", "EN"],
                key="ui_lang",
                label_visibility="collapsed",
                width="content",
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
                try:
                    cb = candidate.stat().st_mtime
                except Exception:
                    cb = None
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
    st.set_page_config(page_title="WECARE 产品解决方案", layout="wide")
    _start_packaged_autoshutdown()
    _render_packaged_quit_button()

    # Debug / emergency: force-clear Streamlit caches via URL
    # Example: https://...streamlit.app/?clear_cache=1
    if _get_query_param_first("clear_cache").strip() in {"1", "true", "yes"}:
        try:
            st.cache_data.clear()
        except Exception:
            pass
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        st.warning("Cache cleared. Please refresh once.")

    # UI 语言：CN / EN（可通过 ?lang=EN 直达）
    lang_from_url = _get_query_param_first("lang").strip().upper()
    if "ui_lang" not in st.session_state:
        st.session_state["ui_lang"] = "EN" if lang_from_url == "EN" else "CN"
    if str(st.session_state.get("ui_lang", "CN")).strip().upper() not in {"CN", "EN"}:
        st.session_state["ui_lang"] = "CN"
    ui_lang = str(st.session_state.get("ui_lang", "CN")).strip().upper() or "CN"

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
        mapping, _meta, main_order, sub_order = load_solution_design(
            str(excel_path), cache_buster
        )
        available_main = [m for m in main_order if m in mapping]
        formula_scenarios = {k: sub_order.get(k, []) for k in available_main}

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
            alias_map = build_scenario_to_solution_title(formula_pptx, str(solutions_pptx_cn), ppt_cache_buster)
    except Exception:
        solutions_deck = {}
        alias_map = {}

    # Header (color matches the selected 功能方向)
    current_cat = _clean_ui_key(st.session_state.get("filter_cat", ""))
    badge_label = current_cat if ui_lang == "CN" else _CATEGORY_LABELS_EN.get(current_cat, current_cat)
    _render_header(series=series, category=current_cat, badge=badge_label)

    scenario_title_en: Dict[str, str] = {}
    scenario_label_en: Dict[str, str] = {}
    pdf_titles_en: Dict[int, str] = {}
    if ui_lang == "EN" and solutions_pdf_en and solutions_pdf_en.exists():
        try:
            pdf_titles_en = load_pdf_solution_titles(
                str(solutions_pdf_en), solutions_pdf_en.stat().st_mtime
            )
        except Exception:
            pdf_titles_en = {}

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

                scenario_idx_en: Dict[str, int] = {}
                if ui_lang == "EN" and solutions_deck and alias_map and (
                    (solutions_pptx_en and solutions_pptx_en.exists()) or pdf_titles_en
                ):
                    try:
                        en_cache_buster = (
                            solutions_pptx_en.stat().st_mtime
                            if solutions_pptx_en and solutions_pptx_en.exists()
                            else None
                        )
                    except Exception:
                        en_cache_buster = None
                    for scen in sub_options:
                        mk = alias_map.get(scen) or alias_map.get(_normalize_match_key(scen)) or scen
                        sol = solutions_deck.get(mk)
                        slide_no = int(sol.get("slide_no", 0)) if sol else 0
                        en_title = ""
                        if slide_no:
                            if solutions_pptx_en and solutions_pptx_en.exists():
                                en_lines = load_pptx_slide_lines(
                                    str(solutions_pptx_en), slide_no, en_cache_buster
                                )
                                en_title = _ppt_solution_title_from_lines(en_lines).strip()
                            if not en_title and pdf_titles_en:
                                en_title = str(pdf_titles_en.get(slide_no, "")).strip()
                        if not en_title:
                            en_title = mk
                        scenario_title_en[scen] = en_title
                        idx = max(1, (slide_no + 1) // 2) if slide_no else 0
                        if idx:
                            scenario_idx_en[scen] = idx
                        scenario_label_en[scen] = f"{idx:02d} · {en_title}" if idx else en_title

                def _format_sub(v: object) -> str:
                    s = str(v)
                    return scenario_label_en.get(s, s) if ui_lang == "EN" else s

                sub_options_sorted = (
                    sorted(sub_options, key=lambda x: scenario_idx_en.get(x, 10**9))
                    if ui_lang == "EN" and scenario_idx_en
                    else sub_options
                )
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

    match_key = alias_map.get(sub) or alias_map.get(_normalize_match_key(sub)) or sub
    ppt_solution = solutions_deck.get(match_key)
    overview_block: Dict[str, object] = {}
    if ppt_solution:
        try:
            overview_lines = list(ppt_solution.get("overview_lines", []))  # type: ignore[arg-type]
            if ui_lang == "EN" and solutions_pptx_en and solutions_pptx_en.exists():
                slide_no = int(ppt_solution.get("slide_no", 1))  # type: ignore[arg-type]
                try:
                    en_cache_buster = solutions_pptx_en.stat().st_mtime
                except Exception:
                    en_cache_buster = None
                en_lines = load_pptx_slide_lines(str(solutions_pptx_en), slide_no, en_cache_buster)
                if en_lines:
                    overview_lines = en_lines
            overview_block = _parse_ppt_overview(overview_lines)
        except Exception:
            overview_block = {}

    overview_info = overview.get(cat, {})
    overview_name = str(overview_info.get("name", "")).strip()
    overview_formula = str(overview_info.get("core_formula", "")).strip()

    with st.container(border=True):
        st.subheader(t("核心配方", "Core Formula"))
        display_name = _ensure_wecpro_registered(overview_name)
        display_formula = overview_formula if ui_lang == "CN" else _to_english_formula(overview_formula)
        if overview_name and overview_formula:
            sep = "：" if ui_lang == "CN" else ":"
            st.markdown(f"**{display_name}**{sep} {display_formula}")
        elif overview_formula:
            st.markdown(display_formula)
        elif overview_name:
            st.markdown(f"**{display_name}**")
        else:
            st.caption(t("（该功能方向暂无‘Sheet2’信息记录）", "(No record found for this health area.)"))

        highlights = [str(x).strip() for x in overview_block.get("highlights", []) if str(x).strip()]  # type: ignore[arg-type]
        if highlights:
            st.markdown(f"**{t('核心功能', 'Core Functions')}**")
            st.markdown(
                "\n".join(f"- {_italicize_microbe_tokens_markdown(x)}" for x in highlights[:4])
            )

    trial_lines = [str(x).strip() for x in overview_block.get("trials", []) if str(x).strip()]  # type: ignore[arg-type]
    trial_entries = _parse_trial_entries(trial_lines)
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
        cap_key = _pick_capsule_scenario(sub, cap_candidates)
        cap_record = capsule_details.get(cat, {}).get(cap_key) if cap_key else None
        capsule_specs = list(cap_record.get("specs", [])) if isinstance(cap_record, dict) else []

    if capsule_specs:
        with st.container(border=True):
            st.subheader(t("规格", "Specifications"))
            clinical_label = t("临床菌配方", "Clinical Strain Formula")
            excipient_label = t("功能性辅料", "Functional Excipients")
            capsule_label = t("胶囊", "Capsule")
            sep = "：" if ui_lang == "CN" else ":"

            clinical_bases: List[str] = []
            clinical_doses: List[str] = []
            for spec in capsule_specs:
                base, dose = _parse_capsule_clinical(spec.get("clinical", ""))
                if base:
                    clinical_bases.append(base)
                if dose:
                    clinical_doses.append(dose)

            base_unique = [b for b in dict.fromkeys(clinical_bases) if b]
            dose_unique = [d for d in dict.fromkeys(clinical_doses) if d]
            if len(base_unique) == 1:
                st.markdown(f"**{clinical_label}**{sep} `{base_unique[0]}`")

            cards: List[str] = []
            for _i, spec in enumerate(capsule_specs[:3]):
                spec_label = str(spec.get("spec", "")).strip()

                m = re.search(r"(?i)(?:Capsule|胶囊)\s*(?P<dose>\d+\s*B)\b", spec_label)
                if m:
                    dose = html.escape(m.group("dose").replace(" ", ""))
                    title_html = (
                        "<div class='spec-title'>"
                        + f"{html.escape(capsule_label)} "
                        + f"<span style='color: var(--accent2);'>{dose}</span>"
                        + "</div>"
                    )
                else:
                    title_html = f"<div class='spec-title'>{html.escape(spec_label)}</div>"

                exc_items_raw = _split_capsule_excipients(str(spec.get("excipients", "")).strip())
                exc_names: List[str] = []
                seen_exc: set[str] = set()
                for raw_item in exc_items_raw:
                    formatted = _strip_mass_units(_format_capsule_excipient_item(raw_item, ui_lang))
                    name_only = _excipient_name_only(formatted)
                    if not name_only or _is_filler_excipient(name_only, ui_lang):
                        continue
                    key = name_only.lower()
                    if key in seen_exc:
                        continue
                    seen_exc.add(key)
                    exc_names.append(name_only)

                if exc_names:
                    exc_html = "".join(f"<div>• {html.escape(n)}</div>" for n in exc_names)
                else:
                    exc_html = "<div>—</div>"

                body = "<div class='spec-box'>" + title_html
                body += f"<div class='spec-meta'>{html.escape(excipient_label)}</div>"
                body += "<div style='font-size:0.92rem; line-height:1.45'>" + exc_html + "</div>"
                body += "</div>"
                cards.append(body)

            if cards:
                grid_html = "<div class='spec-grid'>" + "".join(f"<div>{c}</div>" for c in cards) + "</div>"
                st.markdown(grid_html, unsafe_allow_html=True)

            st.caption(
                t(
                    "实际配方组合可根据客户需求进行定制化设计。",
                    "The actual formulation can be customized according to customer requirements.",
                )
            )

    with st.container(border=True):
        st.markdown(f"### {t('完整解决方案', 'Full Solution')}")
        if not solutions_pptx_cn or not solutions_pptx_cn.exists():
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
        elif not ppt_solution:
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

                slide_no = int(ppt_solution.get("slide_no", 1))  # type: ignore[arg-type]

                render_scale = 2.0
                tool1, tool2, tool3 = st.columns([4, 1, 2])
                with tool1:
                    view_icon = st.segmented_control(
                        t("预览页", "View"),
                        ["←", "⧉", "→"],
                        default="⧉",
                        label_visibility="collapsed",
                        width="content",
                    )
                with tool2:
                    with st.popover(t("显示设置", "View settings"), icon=":material/tune:"):
                        render_scale = st.slider(
                            t("清晰度", "Quality"),
                            min_value=1.0,
                            max_value=3.0,
                            value=2.0,
                            step=0.5,
                        )

                page1 = max(1, slide_no)
                page2 = max(1, slide_no + 1)

                # 下载：始终提供当前 2 页的 PDF
                solution_title = str(match_key or sub)
                if ui_lang == "EN":
                    solution_title = (
                        scenario_title_en.get(sub, "").strip()
                        or str(overview_block.get("title", "")).strip()
                        or solution_title
                    )
                safe_title = _safe_filename_component(solution_title)
                solution_index = max(1, (slide_no + 1) // 2)
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

                with st.spinner(t("正在加载页面...", "Loading pages...")):
                    page_images = render_pdf_pages_png(
                        str(pdf_path),
                        (page1, page2),
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
                    icon = str(view_icon or "⧉")
                    vm = "双页对照" if icon == "⧉" else ("第二页" if icon == "→" else "第一页")

                if page_images and vm == "双页对照":
                    c1, c2 = st.columns(2, gap="large")
                    with c1:
                        _render_pdf_page_card(page_images[0])
                    with c2:
                        _render_pdf_page_card(page_images[1] if len(page_images) > 1 else page_images[0])
                elif page_images and vm == "第二页":
                    _render_pdf_page_card(
                        page_images[1] if len(page_images) > 1 else page_images[0],
                    )
                elif page_images:
                    _render_pdf_page_card(page_images[0])

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
