import importlib.util
import json
import logging
import sys
import unittest
from pathlib import Path


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


if __name__ == "__main__":
    unittest.main()
