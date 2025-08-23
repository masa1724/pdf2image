from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


def images_to_excel(
    img_paths: list[str],
    excel_path: str,
    sheet_name: str = "Sheet1",
    fit_width_px: int = 1200,
    gap_rows: int = 0,
):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # 列幅（概算：1文字幅 ≒ 7.1px）
    approx_chars = max(10, int(fit_width_px / 7.1))
    ws.column_dimensions["A"].width = approx_chars

    row = 1
    for p in img_paths:
        p = str(p)  # 念のため文字列に
        # 元サイズ取得（ファイルは閉じる）
        with Image.open(p) as im:
            orig_w, orig_h = im.size

        # 縮小比率を計算（横幅を fit_width_px に合わせる）
        scale = fit_width_px / orig_w if orig_w else 1.0
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)

        # パスから openpyxl 画像を作成してサイズだけ指定
        ximg = XLImage(p)
        ximg.width = new_w
        ximg.height = new_h

        ws.add_image(ximg, f"A{row}")

        # 行の見積もり（18 , 3 の値は実際に出力して微調整した値で、何かを基準にした値ではない。）
        rows_used = max(1, (new_h // 18) + 1)
        row += rows_used + gap_rows + 3

    wb.save(str(excel_path))
