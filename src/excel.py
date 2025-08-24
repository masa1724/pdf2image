from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage


def images_to_excel(
    img_paths: list[str],
    excel_path: str,
    sheet_name: str = "Sheet1",
    fit_width_px: int = 1000,
    gap_rows: int = 0,
):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # 列幅（概算：1文字幅 ≒ 7.1px）
    approx_chars = max(10, int(fit_width_px / 7.1))
    ws.column_dimensions["A"].width = approx_chars

    # 行の高さを変更（100000は特に深い意味はなく、1画像=100行としてx1000枚入るかなと）
    for r in range(1, 100000):
        ws.row_dimensions[r].height = _px_to_pt(36)  # 36px

    row = 1
    for p in img_paths:
        # 元サイズ取得（ファイルは閉じる）
        with Image.open(p) as im:
            orig_w, orig_h = im.size

        # 縮小比率を計算（横幅を fit_width_px に合わせる）
        scale = fit_width_px / orig_w if orig_w else 1.0
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)

        # パスから openpyxl 画像を作成してサイズ変更
        ximg = XLImage(p)
        ximg.width = new_w
        ximg.height = new_h

        # ワークシートに画像を追加
        ws.add_image(ximg, f"A{row}")

        # 次の画像を貼り付ける行を算出（24は実際に出力しながら調整した値）
        rows_used = max(1, (new_h // 24) + 1)
        row += rows_used + gap_rows

    wb.save(str(excel_path))


def _px_to_pt(px: float, dpi: int = 96) -> float:
    return px * 72.0 / dpi / 1.5  # 1.5は実際に出力しながら調整した値
