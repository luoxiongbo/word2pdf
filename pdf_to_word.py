#!/usr/bin/env python3
"""
PDF -> Word (.docx) converter.

Usage:
  python3 pdf_to_word.py input.pdf
  python3 pdf_to_word.py input.pdf -o output.docx --overwrite
  python3 pdf_to_word.py /path/to/pdf_dir -o /path/to/output_dir --overwrite
"""

import argparse
import io
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
    from docx.shared import Inches
except ImportError:
    Document = None
    Inches = None


def _ensure_dependencies():
    missing = []
    if fitz is None:
        missing.append("PyMuPDF")
    if Document is None or Inches is None:
        missing.append("python-docx")
    if missing:
        raise RuntimeError(
            "Missing dependency: "
            + ", ".join(missing)
            + ". Install with: pip install -r requirements.txt"
        )


def _iter_sorted_blocks(page):
    data = page.get_text("dict")
    blocks = data.get("blocks", [])
    # Sort top-to-bottom, then left-to-right for a stable reading order.
    return sorted(blocks, key=lambda b: (b.get("bbox", [0, 0, 0, 0])[1], b.get("bbox", [0, 0, 0, 0])[0]))


def _extract_text_lines(block):
    lines = []
    for line in block.get("lines", []):
        parts = []
        for span in line.get("spans", []):
            t = (span.get("text") or "").replace("\u00a0", " ")
            if t:
                parts.append(t)
        text = "".join(parts).strip()
        if text:
            lines.append(text)
    return lines


def _pt_to_inches(value_pt):
    return max(0.2, value_pt / 72.0)


def pdf_to_docx(input_pdf: Path, output_docx: Path, keep_page_breaks: bool = True):
    _ensure_dependencies()

    doc = Document()
    pdf = fitz.open(str(input_pdf))

    try:
        for page_idx, page in enumerate(pdf):
            if page_idx > 0 and keep_page_breaks:
                doc.add_page_break()

            blocks = _iter_sorted_blocks(page)
            for block in blocks:
                btype = block.get("type")

                if btype == 0:  # text
                    lines = _extract_text_lines(block)
                    if not lines:
                        continue
                    para = doc.add_paragraph()
                    para.add_run("\n".join(lines))
                    continue

                if btype == 1:  # image
                    image_bytes = block.get("image")
                    bbox = block.get("bbox") or [0, 0, 0, 0]
                    width_pt = max(1.0, float(bbox[2] - bbox[0]))

                    if image_bytes:
                        stream = io.BytesIO(image_bytes)
                        width_inches = min(6.5, _pt_to_inches(width_pt))
                        try:
                            doc.add_picture(stream, width=Inches(width_inches))
                        except Exception:
                            # If image decode fails, skip image but continue text conversion.
                            pass

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_docx))
    finally:
        pdf.close()


def _build_output_path(input_file: Path, output: Path | None):
    if output is None:
        return input_file.with_suffix(".docx")

    if output.exists() and output.is_dir():
        return output / f"{input_file.stem}.docx"

    if input_file.is_file() and output.suffix.lower() == ".docx":
        return output

    # If output doesn't exist and has no .docx suffix, treat as directory.
    if output.suffix.lower() != ".docx":
        return output / f"{input_file.stem}.docx"

    return output


def convert_single(input_file: Path, output: Path | None, overwrite: bool, keep_page_breaks: bool, silent: bool):
    out = _build_output_path(input_file, output)

    if out.exists() and not overwrite:
        if not silent:
            print(f"[skip] exists: {out}")
        return True

    try:
        pdf_to_docx(input_file, out, keep_page_breaks=keep_page_breaks)
        if not silent:
            print(f"[ok] {input_file.name} -> {out}")
        return True
    except Exception as exc:
        if not silent:
            print(f"[error] {input_file.name}: {exc}")
        return False


def convert_dir(input_dir: Path, output_dir: Path | None, overwrite: bool, keep_page_breaks: bool, silent: bool):
    if output_dir is not None and output_dir.suffix.lower() == ".docx":
        raise ValueError("For directory input, --output must be a directory path, not a .docx file")

    files = sorted([p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])
    if not files:
        if not silent:
            print(f"[warn] no .pdf files in: {input_dir}")
        return 0, 0

    out_root = output_dir if output_dir is not None else input_dir
    ok = 0
    fail = 0
    for pdf_file in files:
        result = convert_single(pdf_file, out_root, overwrite, keep_page_breaks, silent)
        if result:
            ok += 1
        else:
            fail += 1
    return ok, fail


def parse_args():
    parser = argparse.ArgumentParser(
        description="Convert PDF to Word (.docx) locally (text-first fidelity)."
    )
    parser.add_argument("input", help="input .pdf file or directory")
    parser.add_argument("-o", "--output", help="output .docx file or output directory")
    parser.add_argument("--overwrite", action="store_true", help="overwrite existing output files")
    parser.add_argument(
        "--no-page-breaks",
        action="store_true",
        help="do not insert page breaks between source PDF pages",
    )
    parser.add_argument("--silent", action="store_true", help="suppress non-error logs")
    return parser.parse_args()


def main():
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve() if args.output else None

    if not input_path.exists():
        print(f"[error] input not found: {input_path}")
        return 1

    keep_page_breaks = not args.no_page_breaks

    if input_path.is_dir():
        ok, fail = convert_dir(
            input_path,
            output_path,
            overwrite=args.overwrite,
            keep_page_breaks=keep_page_breaks,
            silent=args.silent,
        )
        if not args.silent:
            print(f"[summary] ok={ok}, fail={fail}")
        return 0 if fail == 0 else 1

    if input_path.suffix.lower() != ".pdf":
        print("[error] input file must be a .pdf")
        return 1

    success = convert_single(
        input_path,
        output_path,
        overwrite=args.overwrite,
        keep_page_breaks=keep_page_breaks,
        silent=args.silent,
    )
    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
