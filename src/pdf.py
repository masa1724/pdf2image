from pathlib import Path
import io
import fitz  # PyMuPDF
from PIL import Image


def pdf_to_single_image(pdf_path: str, page_nums: list[int], out_path: str, dpi: int):
    """
    指定した PDF ページを、1枚の縦長画像として保存する。

    Args:
        pdf_path: 入力PDFのパス
        page_nums:    1始まりのページ番号リスト（例: [1,2,5]）
        out_path: 出力画像のパス（拡張子でPNG/JPGなど自動判定）
        dpi:      レンダリング解像度（例: 200）
    """

    # DPI(72dpi基準) → 描画倍率へ変換
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    # 指定ページを画像として読み込む
    imgs: list[Image.Image] = []
    with fitz.open(pdf_path) as doc:
        for page_num in page_nums:
            # PyMuPDFは0始まりのため 1引く
            pagePix = doc.load_page(page_num - 1).get_pixmap(matrix=mat, alpha=False)
            # PyMuPDFのバイナリをPillow画像へ（RGBに統一）
            imgs.append(Image.open(io.BytesIO(pagePix.tobytes("png"))).convert("RGB"))

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
    canvas.save(out_path)
