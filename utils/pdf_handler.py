"""
pdf_handler.py
Clean version (no raw images) + fallback-first behavior + combined PDF output.
"""

import os
import argparse
import traceback
from pathlib import Path
from PIL import Image, ImageDraw
import pdfplumber


# -----------------------
# Find matches in PDF
# -----------------------
def find_matches_in_pdf(pdf_path: str, query: str):
    q = query.lower().strip()
    matches = []
    with pdfplumber.open(pdf_path) as pdf:
        for pnum, page in enumerate(pdf.pages):
            words = page.extract_words() or []
            for w in words:
                text = (w.get("text") or "").strip()
                if q in text.lower():
                    matches.append({
                        "page": pnum,
                        "text": text,
                        "x0": float(w["x0"]),
                        "top": float(w["top"]),
                        "x1": float(w["x1"]),
                        "bottom": float(w["bottom"]),
                    })
    return matches


# -----------------------
# Convert PDF page → PIL image
# -----------------------
def render_page(page, resolution=150):
    page_img = page.to_image(resolution=resolution)
    pil = page_img.original
    if pil.mode != "RGB":
        pil = pil.convert("RGB")
    return pil, page_img


# -----------------------
# Draw using pdfplumber.draw_rect
# -----------------------
def draw_with_drawrect(page_img_obj, rects, out_path, stroke="red", stroke_width=3):
    try:
        for r in rects:
            page_img_obj.draw_rect(
                (r["x0"], r["top"], r["x1"], r["bottom"]),
                stroke=stroke,
                stroke_width=stroke_width,
                fill=None
            )
        img = page_img_obj.original.convert("RGB")
        img.save(out_path)
        return True
    except Exception as e:
        print("[draw_rect] failed:", e)
        return False


# -----------------------
# Manual fallback (correct mapping)
# -----------------------
def draw_fallback(pil_img, page_pts, rects, out_path,
                  stroke="red", stroke_width=3):
    img_w, img_h = pil_img.size
    page_w, page_h = page_pts

    scale_x = img_w / page_w
    scale_y = img_h / page_h

    draw = ImageDraw.Draw(pil_img)

    for r in rects:
        x0 = int(r["x0"] * scale_x)
        x1 = int(r["x1"] * scale_x)
        y0 = int(r["top"] * scale_y)
        y1 = int(r["bottom"] * scale_y)

        for w in range(stroke_width):
            draw.rectangle([x0-w, y0-w, x1+w, y1+w], outline=stroke)

    pil_img.save(out_path)
    return True


# -----------------------
# Combine PNGs → single PDF
# -----------------------
def images_to_pdf(image_list, out_pdf):
    imgs = [Image.open(p).convert("RGB") for p in image_list]
    imgs[0].save(out_pdf, save_all=True, append_images=imgs[1:])


# -----------------------
# Main annotation pipeline
# -----------------------
def annotate_pdf_and_build_combined(pdf_path, query,
                                    out_dir="output",
                                    resolution=150,
                                    outline_width=3,
                                    stroke_color="red",
                                    prefer_fallback=True,
                                    force_render=False):

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    matches = find_matches_in_pdf(pdf_path, query)
    print("Matches found:", len(matches))

    matches_by_page = {}
    for m in matches:
        matches_by_page.setdefault(m["page"], []).append(m)

    chosen_images = {}

    with pdfplumber.open(pdf_path) as pdf:
        for pnum, page in enumerate(pdf.pages):

            page_matches = matches_by_page.get(pnum, [])
            if not page_matches and not force_render:
                continue

            pil_img, page_img_obj = render_page(page, resolution)

            fallback_path = out_dir / f"page_{pnum}_annot_fallback.png"
            draw_fallback(
                pil_img.copy(),
                (page.width, page.height),
                page_matches,
                fallback_path,
                stroke=stroke_color,
                stroke_width=outline_width
            )

            chosen_images[pnum] = str(fallback_path)

            drawrect_path = out_dir / f"page_{pnum}_annot_drawrect.png"
            ok = draw_with_drawrect(
                page_img_obj,
                page_matches,
                drawrect_path,
                stroke=stroke_color,
                stroke_width=outline_width
            )
            if ok:
                pass  # just for debugging, fallback remains primary

    ordered = [chosen_images[k] for k in sorted(chosen_images.keys())]
    combined_pdf = out_dir / "annotated_combined.pdf"

    if ordered:
        images_to_pdf(ordered, combined_pdf)
        print("Combined PDF saved at:", combined_pdf)

    return ordered, combined_pdf
    


# -----------------------
# CLI
# -----------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", help="PDF file path")
    parser.add_argument("query", help="Query string")
    parser.add_argument("--outdir", default="output")
    parser.add_argument("--resolution", type=int, default=150)
    parser.add_argument("--outline", type=int, default=3)
    parser.add_argument("--color", default="red")
    parser.add_argument("--prefer-fallback", action="store_true")
    parser.add_argument("--force-render", action="store_true")

    args = parser.parse_args()

    annotate_pdf_and_build_combined(
        args.pdf,
        args.query,
        out_dir=args.outdir,
        resolution=args.resolution,
        outline_width=args.outline,
        stroke_color=args.color,
        prefer_fallback=args.prefer_fallback or True,
        force_render=args.force_render
    )
