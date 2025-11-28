"""
utils/image_handler.py

Image handler for PNG/JPG/JPEG.
Performs OCR → finds text → draws bounding boxes → saves annotated image + PDF.

Requirements:
    pip install pillow pytesseract
    Install Tesseract OCR (Windows):
    https://github.com/UB-Mannheim/tesseract/wiki
"""

import os
import sys
from pathlib import Path
from PIL import Image, ImageDraw
import pytesseract

# Allow import from project root (reuse images_to_pdf)
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    from pdf_handler import images_to_pdf
except:
    raise ImportError("Could not import images_to_pdf from pdf_handler.py. Ensure it exists in project root.")


# ------------------------------
# OCR + bounding boxes
# ------------------------------
def extract_boxes(image_path, query):
    """
    Runs OCR and extracts bounding boxes for matching text.
    Returns: (image_obj, list_of_boxes)
    Each box: (x0, y0, x1, y1)
    """
    img = Image.open(image_path).convert("RGB")
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    boxes = []
    q = query.lower()

    for i in range(len(data["text"])):
        word = data["text"][i].strip()
        if not word:
            continue
        if q in word.lower():   # MATCH
            x = data["left"][i]
            y = data["top"][i]
            w = data["width"][i]
            h = data["height"][i]
            boxes.append((x, y, x + w, y + h))

    return img, boxes


# ------------------------------
# Draw red outline bounding boxes
# ------------------------------
def draw_boxes(img, boxes, out_path, stroke="red", width=4):
    draw = ImageDraw.Draw(img)
    for (x0, y0, x1, y1) in boxes:
        for w in range(width):
            draw.rectangle([x0 - w, y0 - w, x1 + w, y1 + w], outline=stroke)

    img.save(out_path)
    return out_path


# ------------------------------
# Main processor
# ------------------------------
def process_image(image_path, query, out_dir="output",
                  stroke_color="red", outline_width=4):

    image_path = Path(image_path)
    if not image_path.exists():
        raise FileNotFoundError(f"Image not found: {image_path}")

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[image_handler] Input Image : {image_path}")
    print(f"[image_handler] Query       : '{query}'")
    print(f"[image_handler] Output dir  : {out_dir.resolve()}")

    img, boxes = extract_boxes(image_path, query)
    print(f"[image_handler] Matches found: {len(boxes)}")

    out_img_path = out_dir / f"{image_path.stem}_annot.png"
    draw_boxes(img, boxes, out_img_path,
               stroke=stroke_color, width=outline_width)

    print(f"[image_handler] Annotated image saved: {out_img_path}")

    combined_pdf_path = out_dir / "annotated_combined.pdf"
    images_to_pdf([out_img_path], combined_pdf_path)

    print(f"[image_handler] Combined PDF saved: {combined_pdf_path}")

    return str(out_img_path), str(combined_pdf_path)


# ------------------------------
# CLI
# ------------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Image annotation handler")
    parser.add_argument("image", help="Path to image file (.png/.jpg/.jpeg)")
    parser.add_argument("query", help="Search query")
    parser.add_argument("--outdir", default="output", help="Output directory")
    parser.add_argument("--outline", type=int, default=4, help="Bounding box outline width")
    parser.add_argument("--color", default="red", help="Bounding box color")

    args = parser.parse_args()

    try:
        out_img, out_pdf = process_image(
            args.image,
            args.query,
            out_dir=args.outdir,
            stroke_color=args.color,
            outline_width=args.outline
        )
        print("\nOutput Summary:")
        print(" - Annotated Image:", out_img)
        print(" - Combined PDF   :", out_pdf)
    except Exception as e:
        print("[image_handler] ERROR:")
        raise e
