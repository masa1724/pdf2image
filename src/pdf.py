from pathlib import Path
import io
import fitz  # PyMuPDF
from PIL import Image


# 指定したPDFのページを、1枚の縦長画像として保存する。
def pdf_to_single_image(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    dpi: int,
    trim_top_px: int = 0,
    trim_bottom_px: int = 0,
):
    # DPI(72dpi基準) → 描画倍率へ変換
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    # 指定ページを画像として読み込む
    imgs: list[Image.Image] = []
    with fitz.open(pdf_path) as doc:
        for page_num in page_nums:
            # PyMuPDFは0始まりのため 1引く
            page_pix = doc.load_page(page_num - 1).get_pixmap(matrix=mat, alpha=False)
            # PyMuPDFのバイナリをPillow画像へ（RGBに統一）
            im = Image.open(io.BytesIO(page_pix.tobytes("png"))).convert("RGB")

            # ページごとに上下トリム
            top = max(0, min(trim_top_px, im.height))
            bottom = max(0, min(trim_bottom_px, im.height - top))
            new_h = max(1, im.height - top - bottom)
            im = im.crop((0, top, im.width, top + new_h))
            imgs.append(im)

    # 各画像の幅を最大幅に合わせて等比リサイズ
    max_w = max(im.width for im in imgs)
    norm = [
        im
        if im.width == max_w
        else im.resize((max_w, int(im.height * max_w / im.width)))
        for im in imgs
    ]

    # キャンバスを生成（背景は白）
    total_h = sum(im.height for im in norm)
    canvas = Image.new("RGB", (max_w, total_h), (255, 255, 255))

    # キャンバスの上から順に画像を貼り付け
    y = 0
    for im in norm:
        canvas.paste(im, (0, y))
        y += im.height

    # 出力先フォルダが無ければ作成して保存
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    print("out_path={}".format(out_path))
    canvas.save(out_path)
