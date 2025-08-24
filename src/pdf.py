from pathlib import Path
import io
import fitz
from PIL import Image


# 指定した PDF のページ群を 1 枚の縦長画像に結合して保存する。
def pdf_to_single_image(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    trim_top_px: int = 0,
    trim_bottom_px: int = 0,
    dpi: int = 200,
):
    # 画像を読み込む
    imgs: list[Image.Image] = _pdf_to_images(
        pdf_path=pdf_path,
        page_nums=page_nums,
        trim_top_px=trim_top_px,
        trim_bottom_px=trim_bottom_px,
        dpi=dpi,
    )

    # 全画像の内、最も大きい横幅を取得
    max_w = max(img.width for img in imgs)
    # 全画像の高さの合計を算出
    total_h = sum(img.height for img in imgs)

    # キャンバスを生成（背景は白）
    canvas = Image.new("RGB", (max_w, total_h), (255, 255, 255))

    # 出力先フォルダが無ければ作成
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    # キャンバスの上から順に画像を貼り付け
    y = 0
    for img in imgs:
        # 左寄せで貼り付け
        canvas.paste(img, (0, y))
        y += img.height

    # ファイルに書き出す
    canvas.save(out_path, optimize=True)


# 指定した PDF のページ群を 1ページ=1画像 で個別保存し、そのパス一覧を返す。
def pdf_to_images(
    pdf_path: str,
    page_nums: list[int],
    out_path: str,
    trim_top_px: int = 0,
    trim_bottom_px: int = 0,
    dpi: int = 200,
):
    # 画像を読み込む
    imgs: list[Image.Image] = _pdf_to_images(
        pdf_path=pdf_path,
        page_nums=page_nums,
        trim_top_px=trim_top_px,
        trim_bottom_px=trim_bottom_px,
        dpi=dpi,
    )

    # 出力先フォルダが無ければ作成
    path_obj = Path(out_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    # ファイル名の連番の桁数を要素数に合わせる
    digits = len(str(len(imgs)))

    img_file_paths: list[str] = []

    for idx, img in enumerate(imgs, start=1):
        # キャンバスを生成（背景は白）
        canvas = Image.new("RGB", (img.width, img.height), (255, 255, 255))

        # 左寄せで貼り付け
        canvas.paste(img, (0, 0))

        # ファイルに書き出す
        new_path = path_obj.with_name(
            f"{path_obj.stem}_page{page_nums[idx - 1]:0{digits}d}{path_obj.suffix}"
        )
        canvas.save(new_path, optimize=True)

        img_file_paths.append(new_path)

    return img_file_paths


# 指定した PDF のページ群を画像として読み込み、その一覧を返す。
def _pdf_to_images(
    pdf_path: str,
    page_nums: list[int],
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
            # 1ページ分を読み込み
            page_pix = doc.load_page(page_num - 1).get_pixmap(matrix=mat, alpha=False)
            # PyMuPDFのバイナリをPillow画像に変換
            img = Image.open(io.BytesIO(page_pix.tobytes("png"))).convert("RGB")

            # ページごとに上下トリム
            top = max(0, min(trim_top_px, img.height))
            bottom = max(0, min(trim_bottom_px, img.height - top))
            new_h = max(1, img.height - top - bottom)
            img = img.crop((0, top, img.width, top + new_h))

            imgs.append(img)

    return imgs
