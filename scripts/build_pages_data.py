#!/usr/bin/env python3
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List


def _unwrap(fn):
    return getattr(fn, "__wrapped__", fn)


def _safe_str(v: object) -> str:
    return str(v or "").strip()


def main() -> int:
    repo = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(repo))
    docs = repo / "docs"
    out_json = docs / "data" / "solutions.json"

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

    formula_scenarios: Dict[str, List[str]] = _unwrap(app.load_formula_scenarios)(
        str(formula_pptx), None
    )
    alias_map: Dict[str, str] = _unwrap(app.build_scenario_to_solution_title)(
        str(formula_pptx), str(solutions_pptx_cn), None
    )
    solutions_deck: Dict[str, Dict[str, Any]] = _unwrap(app.load_ppt_solution_deck)(
        str(solutions_pptx_cn), None
    )

    def match_key(scen: str) -> str:
        n = _unwrap(app._normalize_match_key)(scen)
        return _safe_str(alias_map.get(scen) or alias_map.get(n) or scen)

    categories: List[Dict[str, Any]] = []
    for cat, scenarios in formula_scenarios.items():
        theme = app._CATEGORY_THEME.get(cat, {})
        accent1 = theme.get("accent1", "#6366F1")
        accent2 = theme.get("accent2", "#EC4899")

        cat_obj: Dict[str, Any] = {
            "key": cat,
            "label": {"CN": cat, "EN": app._CATEGORY_LABELS_EN.get(cat, cat)},
            "accent1": accent1,
            "accent2": accent2,
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

            cat_obj["scenarios"].append(
                {
                    "key": scen,
                    "label": {"CN": scen, "EN": f"{idx:02d} · {en_title}"},
                    "title": {"CN": cn_title, "EN": en_title},
                    "index": idx,
                    "page1": slide_no,
                    "page2": slide_no + 1,
                }
            )

        if cat_obj["scenarios"]:
            categories.append(cat_obj)

    out: Dict[str, Any] = {
        "pdf": {
            "CN": "./assets/solutions_cn.pdf",
            "EN": "./assets/solutions_en.pdf",
        },
        "categories": categories,
    }

    out_json.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {out_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
