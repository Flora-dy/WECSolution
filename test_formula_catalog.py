import importlib.util
import json
import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parent
logging.disable(logging.CRITICAL)
SPEC = importlib.util.spec_from_file_location("wecsolution_app", ROOT / "app.py")
app = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = app
SPEC.loader.exec_module(app)


class FormulaCatalogTest(unittest.TestCase):
    def test_gastrointestinal_formula_renders_only_gihealth805(self):
        for language in ("CN", "EN"):
            with self.subTest(language=language):
                rendered = app._render_formula_variants_html("胃肠健康", language)
                self.assertIn("WecPro®-GIHealth805", rendered)
                self.assertEqual(rendered.count("class='v-box'"), 1)
                self.assertNotIn("WecPro®-GUT99", rendered)
                self.assertNotIn("WecPro®-DigestBi", rendered)
                self.assertNotIn("WecPro®-Pyloclear", rendered)

    def test_static_catalog_exposes_the_same_single_formula(self):
        catalog = json.loads(
            (ROOT / "docs/data/pages_data.json").read_text(encoding="utf-8")
        )
        item = next(
            row
            for row in catalog["formula"]["items"]
            if row["direction"] == "胃肠健康"
        )

        products = [variant["product"]["EN"] for variant in item["variants"]]
        self.assertEqual(products, ["WecPro®-GIHealth805"])
        self.assertEqual(item["product"], {"CN": "1个配方", "EN": "1 Formula"})

    def test_single_gastrointestinal_formula_uses_standard_full_width_details(self):
        item = {
            "direction": "胃肠健康",
            "product": "WecPro®-GIHealth805",
            "benefit": "调节胃肠健康",
            "strains": ["BLa80", "LRa05"],
        }
        rendered_blocks = []

        with (
            patch.object(app.st, "session_state", {"ui_lang": "EN"}),
            patch.object(app, "load_wecpro_formula_catalog", return_value=[item]),
            patch.object(
                app.st,
                "markdown",
                side_effect=lambda body, **_: rendered_blocks.append(body),
            ),
        ):
            app._render_wecpro_formula_page()

        gastrointestinal = next(
            block for block in rendered_blocks if "Gastrointestinal Health" in block
        )
        self.assertIn("<div class='kv-table'>", gastrointestinal)
        self.assertNotIn("<div class='v-grid'>", gastrointestinal)
        self.assertIn("Benefits", gastrointestinal)
        self.assertIn("Core Formula", gastrointestinal)
        self.assertIn(
            "Supports gastrointestinal health, helps relieve constipation and diarrhea",
            gastrointestinal,
        )
        self.assertNotIn("Premium / Base / Active Probiotic Yogurt", gastrointestinal)


if __name__ == "__main__":
    unittest.main()
