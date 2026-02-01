#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path


def render(pdf_path: Path, out_dir: Path, scale: float) -> int:
    import fitz  # PyMuPDF

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    try:
        total = int(getattr(doc, "page_count", len(doc)))
        for i in range(total):
            page_no = i + 1
            out = out_dir / f"{page_no:03d}.png"
            if out.exists():
                continue
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            out.write_bytes(pix.tobytes("png"))
        return total
    finally:
        doc.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--cn", required=True, help="CN PDF path")
    ap.add_argument("--en", required=True, help="EN PDF path")
    ap.add_argument("--out", required=True, help="Output base dir (will create CN/EN)")
    ap.add_argument("--scale", type=float, default=1.6)
    args = ap.parse_args()

    base = Path(args.out)
    cn_pdf = Path(args.cn)
    en_pdf = Path(args.en)
    if not cn_pdf.exists():
        raise SystemExit(f"Missing: {cn_pdf}")
    if not en_pdf.exists():
        raise SystemExit(f"Missing: {en_pdf}")

    total_cn = render(cn_pdf, base / "CN", float(args.scale))
    total_en = render(en_pdf, base / "EN", float(args.scale))
    print(f"Rendered CN: {total_cn} pages -> {base/'CN'}")
    print(f"Rendered EN: {total_en} pages -> {base/'EN'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

