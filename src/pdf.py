from pathlib import Path
import io
import fitz
from PIL import Image


# 指定したPDFのページを、1枚の縦長画像として保存する。
def pdf_to_single_image(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    trim_top_px: int = 0,
    trim_bottom_px: int = 0,
    dpi: int = 200,
):
    # 指定ページを画像として読み込む
    imgs: list[Image.Image] = __pdf_to_images(
        pdf_path=pdf_path,
        page_nums=page_nums,
        out_path=out_path,
        trim_top_px=trim_top_px,
        trim_bottom_px=trim_bottom_px,
        dpi=dpi,
    )

    # 全画像の内、最も大きい横幅を取得
    max_w = max(im.width for im in imgs)
    # 全画像の高さの合計値を算出
    total_h = sum(im.height for im in imgs)

    # キャンバスを生成（背景は白）
    canvas = Image.new("RGB", (max_w, total_h), (255, 255, 255))

    # 出力先フォルダが無ければ作成して保存
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    # キャンバスの上から順に画像を貼り付け
    y = 0
    for im in imgs:
        canvas.paste(im, (0, y))
        y += im.height

    # ファイルに書き出す
    canvas.save(out_path)


# 指定したPDFのページを、1ページずつ画像変換し、そのファイルパスを返す。
def pdf_to_images(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    trim_top_px: int = 0,
    trim_bottom_px: int = 0,
    dpi: int = 200,
):
    # 指定ページを画像として読み込む
    imgs: list[Image.Image] = __pdf_to_images(
        pdf_path=pdf_path,
        page_nums=page_nums,
        out_path=out_path,
        trim_top_px=trim_top_px,
        trim_bottom_px=trim_bottom_px,
        dpi=dpi,
    )

    # 出力先フォルダが無ければ作成して保存
    path_obj = Path(out_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    # 連番の桁数を決定（要素数に応じてゼロ埋め）
    digits = len(str(len(imgs)))

    img_file_paths: list[str] = []

    for idx, im in enumerate(imgs, start=1):
        # キャンバスを生成（背景は白）
        canvas = Image.new("RGB", (im.width, im.height), (255, 255, 255))

        # キャンバスの上に画像を貼り付け
        canvas.paste(im, (0, 0))

        # ファイルに書き出す
        new_path = path_obj.with_name(
            f"{path_obj.stem}_part{idx:0{digits}d}{path_obj.suffix}"
        )
        canvas.save(new_path)
        img_file_paths.append(new_path)

    return img_file_paths


# 指定したPDFのページを、1枚の縦長画像として保存する。
def __pdf_to_images(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    trim_top_px: int,
    trim_bottom_px: int,
    dpi: int,
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

    return imgs
