#!/usr/bin/env python3
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple


def _unwrap(fn):
    return getattr(fn, "__wrapped__", fn)


def _safe_str(v: object) -> str:
    return str(v or "").strip()


def main() -> int:
    repo = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(repo))
    docs = repo / "docs"
    out_json = docs / "data" / "pages_data.json"

    # Import app.py (Streamlit UI) only for its PPTX parsing helpers.
    import app  # type: ignore

    formula_pptx = repo / "Final" / "Formula&Solution.pptx"
    solutions_pptx_cn = repo / "Final" / "43 Solutions解决方案中文版20260130.pptx"
    solutions_pptx_en = repo / "Final" / "43 Solutions解决方案英文版20260130.pptx"

    if not formula_pptx.exists():
        raise SystemExit(f"Missing: {formula_pptx}")
    if not solutions_pptx_cn.exists():
        raise SystemExit(f"Missing: {solutions_pptx_cn}")
    if not solutions_pptx_en.exists():
        raise SystemExit(f"Missing: {solutions_pptx_en}")

    # Shared: scenarios + solutions deck
    formula_scenarios: Dict[str, List[str]] = _unwrap(app.load_formula_scenarios)(str(formula_pptx), None)
    alias_map: Dict[str, str] = _unwrap(app.build_scenario_to_solution_title)(str(formula_pptx), str(solutions_pptx_cn), None)
    solutions_deck: Dict[str, Dict[str, Any]] = _unwrap(app.load_ppt_solution_deck)(str(solutions_pptx_cn), None)

    # Excel sheet2 core formula per category
    excel = repo / "产品配方设计最新.xlsx"
    overview: Dict[str, Any] = {}
    if excel.exists():
        try:
            overview = _unwrap(app.load_product_overview)(str(excel), None)
        except Exception:
            overview = {}

    def match_key(scen: str) -> str:
        n = _unwrap(app._normalize_match_key)(scen)
        return _safe_str(alias_map.get(scen) or alias_map.get(n) or scen)

    # Clinical article links
    clinical_links: Dict[str, str] = {}
    clinical_xlsx = repo / "Final" / "Clinicaldata0201.xlsx"
    if clinical_xlsx.exists():
        try:
            clinical_links = _unwrap(app.load_clinical_article_links)(str(clinical_xlsx), None)
        except Exception:
            clinical_links = {}

    # Capsule specs per (category, scenario) from CN/EN sheets
    capsule_specs_cn_by_cat: Dict[str, Any] = {}
    capsule_specs_en_by_cat: Dict[str, Any] = {}
    capsule_xlsx = repo / "Final" / "Capsule配方详情.xlsx"
    if capsule_xlsx.exists():
        try:
            capsule_specs_cn_by_cat = _unwrap(app.load_capsule_details)(str(capsule_xlsx), "CN", None)
        except Exception:
            capsule_specs_cn_by_cat = {}
        try:
            capsule_specs_en_by_cat = _unwrap(app.load_capsule_details)(str(capsule_xlsx), "EN", None)
        except Exception:
            capsule_specs_en_by_cat = {}

    categories: List[Dict[str, Any]] = []
    for cat, scenarios in formula_scenarios.items():
        theme = app._CATEGORY_THEME.get(cat, {})
        accent1 = theme.get("accent1", "#6366F1")
        accent2 = theme.get("accent2", "#EC4899")

        overview_info = overview.get(cat, {}) if isinstance(overview, dict) else {}
        core_name = _safe_str(overview_info.get("name", ""))
        core_formula_cn = _safe_str(overview_info.get("core_formula", ""))

        core = {
            "name": {"CN": _unwrap(app._ensure_wecpro_registered)(core_name) if core_name else "", "EN": _unwrap(app._ensure_wecpro_registered)(core_name) if core_name else ""},
            "formula": {"CN": core_formula_cn, "EN": _unwrap(app._to_english_formula)(core_formula_cn) if core_formula_cn else ""},
        }

        cat_obj: Dict[str, Any] = {
            "key": cat,
            "label": {"CN": cat, "EN": app._CATEGORY_LABELS_EN.get(cat, cat)},
            "accent1": accent1,
            "accent2": accent2,
            "core": core,
            "scenarios": [],
        }

        for scen in scenarios:
            mk = match_key(scen)
            sol = solutions_deck.get(mk) or solutions_deck.get(scen)
            slide_no = int(sol.get("slide_no", 0)) if sol else 0
            if not slide_no:
                continue

            cn_lines = _unwrap(app.load_pptx_slide_lines)(str(solutions_pptx_cn), slide_no, None)
            cn_title = _safe_str(_unwrap(app._ppt_solution_title_from_lines)(cn_lines)) or mk

            en_lines = _unwrap(app.load_pptx_slide_lines)(str(solutions_pptx_en), slide_no, None)
            en_title = _safe_str(_unwrap(app._ppt_solution_title_from_lines)(en_lines)) or mk

            idx = max(1, (slide_no + 1) // 2)

            # Highlights + clinical registrations are extracted from the overview slide text
            highlights_cn: List[str] = []
            highlights_en: List[str] = []
            clinical_rows: List[Dict[str, Any]] = []
            try:
                ob_cn = _unwrap(app._parse_ppt_overview)(cn_lines)
                ob_en = _unwrap(app._parse_ppt_overview)(en_lines)
                highlights_cn = [str(x).strip() for x in ob_cn.get("highlights", []) if str(x).strip()]
                highlights_en = [str(x).strip() for x in ob_en.get("highlights", []) if str(x).strip()]
                trials_cn = [str(x).strip() for x in ob_cn.get("trials", []) if str(x).strip()]
                trial_entries = _unwrap(app._parse_trial_entries)(trials_cn)
                for key, ids in trial_entries:
                    badges: List[str] = []
                    for reg_id in ids:
                        rid = (reg_id or "").strip().replace(" ", "")
                        url = _safe_str(clinical_links.get(rid, ""))
                        if url:
                            badges.append(
                                "<a class='tile-badge tile-badge-link' "
                                f"href='{url}' target='_blank' rel='noopener noreferrer'>"
                                f"{reg_id}</a>"
                            )
                        else:
                            badges.append(f"<span class='tile-badge'>{reg_id}</span>")
                    clinical_rows.append(
                        {
                            "key": {"CN": key, "EN": key},
                            "ids": badges,
                        }
                    )
            except Exception:
                pass

            # Specs: choose scenario record by matching label
            specs_out: List[Dict[str, Any]] = []
            try:
                cat_specs_cn = (
                    capsule_specs_cn_by_cat.get(cat, {}) if isinstance(capsule_specs_cn_by_cat, dict) else {}
                )
                cat_specs_en = (
                    capsule_specs_en_by_cat.get(cat, {}) if isinstance(capsule_specs_en_by_cat, dict) else {}
                )
                cap_candidates = list(cat_specs_cn.keys()) if isinstance(cat_specs_cn, dict) else []
                cap_key = _unwrap(app._pick_capsule_scenario)(scen, cap_candidates)
                cap_record_cn = cat_specs_cn.get(cap_key) if isinstance(cat_specs_cn, dict) and cap_key else None
                cap_record_en = cat_specs_en.get(cap_key) if isinstance(cat_specs_en, dict) and cap_key else None

                specs_cn = list(cap_record_cn.get("specs", [])) if isinstance(cap_record_cn, dict) else []
                specs_en = list(cap_record_en.get("specs", [])) if isinstance(cap_record_en, dict) else []

                def normalize_total(t: str) -> str:
                    return _unwrap(app._strip_mass_units)(
                        _unwrap(app._normalize_total_text)(_safe_str(t))
                    )

                def parse_excipients(raw: str, lang: str) -> List[str]:
                    parts = _unwrap(app._split_capsule_excipients)(_safe_str(raw))
                    items: List[str] = []
                    for p in parts:
                        x = _unwrap(app._format_capsule_excipient_item)(p, lang)
                        x = _unwrap(app._strip_mass_units)(_safe_str(x))
                        if x:
                            items.append(x)
                    return items

                def title_from_spec(spec_label: str, lang: str) -> str:
                    s = _safe_str(spec_label)
                    if not s:
                        return ""
                    # Prefer "Capsule 120B" style title
                    import re

                    m = re.search(r"(?i)(?:Capsule|胶囊)\s*(?P<dose>\d+\\s*B)\\b", s)
                    if m:
                        dose = m.group("dose").replace(" ", "")
                        return f"{'Capsule' if lang == 'EN' else '胶囊'} {dose}"
                    return s

                # Keep the first 3 cards like Streamlit
                for i in range(min(3, max(len(specs_cn), len(specs_en)))):
                    s_cn = specs_cn[i] if i < len(specs_cn) else {}
                    s_en = specs_en[i] if i < len(specs_en) else {}

                    spec_label_cn = _safe_str(s_cn.get("spec", "")) or _safe_str(s_en.get("spec", ""))
                    spec_label_en = _safe_str(s_en.get("spec", "")) or _safe_str(s_cn.get("spec", ""))

                    clinical_cn = _safe_str(s_cn.get("clinical", ""))
                    clinical_en = _safe_str(s_en.get("clinical", "")) or clinical_cn

                    exc_cn = parse_excipients(_safe_str(s_cn.get("excipients", "")), "CN")
                    exc_en = parse_excipients(_safe_str(s_en.get("excipients", "")), "EN")

                    total_cn = normalize_total(_safe_str(s_cn.get("total", "")))
                    total_en = normalize_total(_safe_str(s_en.get("total", "")) or _safe_str(s_cn.get("total", "")))

                    specs_out.append(
                        {
                            "title": {"CN": title_from_spec(spec_label_cn, "CN"), "EN": title_from_spec(spec_label_en, "EN")},
                            "clinical": {"CN": clinical_cn, "EN": clinical_en},
                            "excipients": {"CN": exc_cn, "EN": exc_en},
                            "total": {"CN": total_cn, "EN": total_en},
                        }
                    )
            except Exception:
                specs_out = []

            cat_obj["scenarios"].append(
                {
                    "key": scen,
                    "label": {"CN": scen, "EN": f"{idx:02d} · {en_title}"},
                    "title": {"CN": cn_title, "EN": en_title},
                    "index": idx,
                    "page1": slide_no,
                    "page2": slide_no + 1,
                    "highlights": {"CN": highlights_cn, "EN": highlights_en},
                    "clinical": clinical_rows,
                    "specs": specs_out,
                }
            )

        if cat_obj["scenarios"]:
            categories.append(cat_obj)

    # WecLac strains from PPTX (CN/EN)
    weclac_cn = _unwrap(app.load_weclac_catalog)(str(repo / "Final" / "WecLac.pptx"), "CN", None)
    weclac_en = _unwrap(app.load_weclac_catalog)(str(repo / "Final" / "WecLac.pptx"), "EN", None)
    strains_cn = list(weclac_cn.get("strains", [])) if isinstance(weclac_cn, dict) else []
    strains_en = list(weclac_en.get("strains", [])) if isinstance(weclac_en, dict) else []

    code_to_en: Dict[str, Dict[str, Any]] = {}
    for it in strains_en:
        name = _safe_str(it.get("name", ""))
        base_name, code = _unwrap(app._extract_strain_code)(name)
        if code:
            code_to_en[code] = {"name": name, "base_name": base_name, **it}

    # Enrich CN list to 12
    weclac_items: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for it in strains_cn:
        name = _safe_str(it.get("name", ""))
        base_name, code = _unwrap(app._extract_strain_code)(name)
        if not code or code in seen:
            continue
        seen.add(code)
        en_it = code_to_en.get(code, {})
        sci = app._STRAIN_SCI_NAMES.get(code, "")
        latin_html = _unwrap(app._format_sci_name_html)(sci) if sci else ""
        icon_path = repo / "docs" / "assets" / "strains" / f"{code}.png"
        icon = f"./assets/strains/{code}.png" if icon_path.exists() else ""
        weclac_items.append(
            {
                "code": code,
                "base_name": {"CN": base_name, "EN": _safe_str(en_it.get("base_name", ""))},
                "latin_html": latin_html,
                "icon": icon,
                "feature": {"CN": _safe_str(it.get("feature", "")), "EN": _safe_str(en_it.get("feature", ""))},
                "clinical": {"CN": _safe_str(it.get("clinical", "")), "EN": _safe_str(en_it.get("clinical", ""))},
                "patent": {"CN": _safe_str(it.get("patent", "")), "EN": _safe_str(en_it.get("patent", ""))},
                "spec": {"CN": _safe_str(it.get("spec", "")), "EN": _safe_str(en_it.get("spec", ""))},
            }
        )
        if len(weclac_items) >= 12:
            break

    def first_functions(strains: List[Dict[str, Any]]) -> List[Dict[str, str]]:
        for it in strains:
            f = it.get("functions", [])
            if isinstance(f, list) and f:
                return [x for x in f if isinstance(x, dict)]
        return []

    functions_cn = first_functions(strains_cn)
    functions_en = first_functions(strains_en)

    # Align EN area labels to match Streamlit mapping
    area_aliases = {
        "Emotional & Cognitive Health": "Mental Health",
        "Emotional and Cognitive Health": "Mental Health",
        "Oral Health": "Dental & Oral Health",
        "Immune Health": "Immunological Health",
        "Infant and Child Health": "Infant Health",
        "Infant & Child Health": "Infant Health",
    }

    def map_area_en(label: str) -> str:
        raw = (label or "").strip()
        if not raw:
            return ""
        for k, v in area_aliases.items():
            if raw.lower() == k.lower():
                return v
        return raw

    areas: List[str] = []
    for f in functions_cn:
        d = _safe_str(f.get("direction", ""))
        if d:
            areas.append(d)
    areas = list(dict.fromkeys(areas))

    # Build area detail rows: direction -> desc (CN/EN)
    cn_map = { _safe_str(f.get("direction","")): _safe_str(f.get("desc","")) for f in functions_cn if _safe_str(f.get("direction","")) }
    en_map = { map_area_en(_safe_str(f.get("direction",""))): _safe_str(f.get("desc","")) for f in functions_en if _safe_str(f.get("direction","")) }
    area_details: List[Dict[str, Any]] = []
    for d_cn in areas:
        d_en = app._CATEGORY_LABELS_EN.get(d_cn, "") or map_area_en(d_cn) or d_cn
        desc_cn = cn_map.get(d_cn, "")
        desc_en = en_map.get(d_en, "") or ""
        area_details.append(
            {
                "direction": {"CN": d_cn, "EN": d_en},
                "desc_html": {
                    "CN": _unwrap(app._italicize_microbe_tokens_html)(desc_cn) if desc_cn else "",
                    "EN": _unwrap(app._italicize_microbe_tokens_html)(desc_en) if desc_en else "",
                },
            }
        )

    # WecPro® Formula list from Formula.pptx
    formula_items_raw = _unwrap(app.load_wecpro_formula_catalog)(str(repo / "Final" / "Formula.pptx"), None)
    order = [d for d in app._FORMULA_SLIDE_TO_DIRECTION.values() if d]
    direction_to_item: Dict[str, Dict[str, Any]] = {}
    for it in formula_items_raw:
        d = _safe_str(it.get("direction", ""))
        if d:
            direction_to_item[d] = it

    formula_items: List[Dict[str, Any]] = []
    for direction in order[:7]:
        it = direction_to_item.get(direction, {})
        product = _safe_str(it.get("product", ""))
        benefit = _safe_str(it.get("benefit", ""))
        strains = [str(s).strip() for s in it.get("strains", []) if str(s).strip()] if isinstance(it.get("strains", []), list) else []
        if strains:
            strains_text_cn = "、".join(strains)
        else:
            strains_text_cn = ""
        benefit_en = app._WECPRO_FORMULA_BENEFIT_EN.get(direction, benefit)
        direction_en = app._CATEGORY_LABELS_EN.get(direction, direction)

        # For EN formula, show codes + italic sci names where known
        strains_html_en_parts: List[str] = []
        for line in strains:
            codes = _unwrap(app._extract_strain_codes)(line)
            for code in codes:
                sci = app._STRAIN_SCI_NAMES.get(code, "")
                if sci:
                    strains_html_en_parts.append(f"<div>{_unwrap(app._format_sci_name_html)(sci)} {code}</div>")
                else:
                    strains_html_en_parts.append(f"<div>{code}</div>")
        strains_html_en = "".join(strains_html_en_parts)

        formula_items.append(
            {
                "direction": direction,
                "direction_label": {"CN": direction, "EN": direction_en},
                "product": {"CN": product, "EN": product},
                "benefit": {"CN": benefit, "EN": benefit_en},
                "core_formula": {"CN": strains_text_cn, "EN": strains_html_en},
            }
        )

    out: Dict[str, Any] = {
        "weclac": {
            "strains": weclac_items,
            "core_codes": sorted(list(getattr(app, "WECLAC_CORE_CODES", set()))),
            "areas": areas,
            "area_details": area_details,
        },
        "formula": {"items": formula_items},
        "solutions": {
            "pdf": {"CN": "./assets/solutions_cn.pdf", "EN": "./assets/solutions_en.pdf"},
            "categories": categories,
        },
    }

    out_json.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {out_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
