from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pdf import pdf_to_single_image, pdf_to_images
from excel import images_to_excel
import fitz
import traceback


class App(tk.Tk):
    # GUIの構築
    def __init__(self):
        super().__init__()

        # タイトル
        self.title("PDF画像変換")
        # ウィンドウサイズ（横x縦)
        self.geometry("520x300")

        self.pdf_path: Path | None = None
        self.out_dir_path: Path | None = None
        self.page_count = 0

        pad = {"padx": 8, "pady": 6}

        # 行0: PDFを選択
        self.btn_pick = ttk.Button(self, text="PDFを選択", command=self.cmd_select_pdf)
        self.btn_pick.grid(row=0, column=0, **pad, sticky="w")

        self.lbl_pdf = ttk.Label(
            self, text="（未選択）", anchor="w", justify="left", wraplength=440
        )
        self.lbl_pdf.grid(row=0, column=1, columnspan=3, **pad, sticky="ew")

        # 行1: ページ範囲（開始/終了）
        ttk.Label(self, text="開始ページ").grid(row=1, column=0, **pad, sticky="w")
        self.var_page_start = tk.StringVar(value="1")
        self.spn_page_start = ttk.Spinbox(
            self, from_=1, to=1, width=8, textvariable=self.var_page_start
        )
        self.spn_page_start.grid(row=1, column=1, **pad, sticky="w")

        ttk.Label(self, text="終了ページ").grid(row=1, column=2, **pad, sticky="w")
        self.var_page_end = tk.StringVar(value="1")
        self.spn_page_end = ttk.Spinbox(
            self, from_=1, to=1, width=8, textvariable=self.var_page_end
        )
        self.spn_page_end.grid(row=1, column=3, **pad, sticky="w")

        # 行2: トリム量（上/下, px）
        ttk.Label(self, text="上トリム(px)").grid(row=2, column=0, **pad, sticky="w")
        self.var_trim_top = tk.StringVar(value="0")
        self.spn_trim_top = ttk.Spinbox(
            self, from_=0, to=99999, width=8, textvariable=self.var_trim_top
        )
        self.spn_trim_top.grid(row=2, column=1, **pad, sticky="w")

        ttk.Label(self, text="下トリム(px)").grid(row=2, column=2, **pad, sticky="w")
        self.var_trim_bottom = tk.StringVar(value="0")
        self.spn_trim_bottom = ttk.Spinbox(
            self, from_=0, to=99999, width=8, textvariable=self.var_trim_bottom
        )
        self.spn_trim_bottom.grid(row=2, column=3, **pad, sticky="w")

        # 行3: 保存先
        self.btn_out = ttk.Button(
            self, text="保存先を選択", command=self.cmd_select_out_dir
        )
        self.btn_out.grid(row=3, column=0, **pad, sticky="w")

        self.lbl_out_dir = ttk.Label(
            self, text="（未選択）", anchor="w", justify="left", wraplength=440
        )
        self.lbl_out_dir.grid(row=3, column=1, columnspan=3, **pad, sticky="ew")

        # 行4: 出力モード（ラジオボタン）
        ttk.Label(self, text="出力").grid(row=4, column=0, **pad, sticky="w")
        self.var_mode = tk.StringVar(value="image")  # "image" or "excel"
        ttk.Radiobutton(
            self,
            text="1枚の画像",
            value="image",
            variable=self.var_mode,
        ).grid(row=4, column=1, **pad, sticky="w")
        ttk.Radiobutton(
            self,
            text="Excelに1ページずつ貼り付け",
            value="excel",
            variable=self.var_mode,
        ).grid(row=4, column=2, columnspan=2, **pad, sticky="w")

        # 行5: 実行
        self.btn_run = ttk.Button(self, text="実行", command=self.cmd_run)
        self.btn_run.grid(row=5, column=3, **pad, sticky="e")

        # 行6: ステータス
        self.lbl_status = ttk.Message(
            self, text="", anchor="w", justify="left", wraplength=440
        )
        self.lbl_status.grid(row=6, column=0, columnspan=3, **pad, sticky="ew")

        for i in range(4):
            self.grid_columnconfigure(i, weight=0)
        self.grid_columnconfigure(1, weight=1)

    # 「PDFを選択」押下時の処理
    def cmd_select_pdf(self):
        path = filedialog.askopenfilename(
            title="PDFを選択", filetypes=[("PDF", "*.pdf"), ("All Files", "*.*")]
        )
        if not path:
            return
        self.pdf_path = Path(path)
        self.lbl_pdf.config(text=str(self.pdf_path))

        # 選択したPDFのページ数を取得
        try:
            with fitz.open(self.pdf_path) as doc:
                self.page_count = len(doc)
        except Exception as e:
            messagebox.showerror("エラー", f"PDFを開けませんでした。\n{e}")
            return

        # 選択したPDFのページ数に従い、入力値の下限・上限を更新
        self.lbl_pdf.config(text=f"{self.pdf_path.name}（{self.page_count}ページ）")
        for spn in (self.spn_page_start, self.spn_page_end):
            spn.config(state="normal")
            spn.config(from_=1, to=self.page_count)

        # 開始ぺージを1、終了ぺージを選択したPDFのページ数に更新
        self.var_page_start.set("1")
        self.var_page_end.set(str(self.page_count))

    # 「保存先を選択」押下時の処理
    def cmd_select_out_dir(self):
        path = filedialog.askdirectory(title="保存先フォルダを選択")
        if not path:
            return
        self.out_dir_path = Path(path)
        self.lbl_out_dir.config(text=str(self.out_dir_path))

    # 「実行」押下時の処理
    def cmd_run(self):
        if not self.pdf_path:
            messagebox.showwarning("警告", "PDF が選択されていません。")
            return

        if not self.out_dir_path:
            messagebox.showwarning("警告", "保存先 が選択されていません。")
            return

        try:
            page_start = int(self.spn_page_start.get())
            page_end = int(self.spn_page_end.get())
        except ValueError:
            messagebox.showwarning("警告", "ページ番号は数値で入力してください。")
            return
        if not (
            1 <= page_start <= self.page_count and 1 <= page_end <= self.page_count
        ):
            messagebox.showwarning("警告", "範囲外のページが指定されています。")
            return
        if page_start > page_end:
            messagebox.showwarning("警告", "開始ページが終了ページを超えています。")
            return

        try:
            trim_top = int(self.spn_trim_top.get())
            trim_bottom = int(self.spn_trim_bottom.get())
        except ValueError:
            messagebox.showwarning("警告", "トリム量は数値で入力してください。")
            return
        if trim_top < 0 or trim_bottom < 0:
            messagebox.showwarning("警告", "トリム量は0以上で入力してください。")
            return

        # 実行
        try:
            self.btn_run.config(state="disabled")
            self.lbl_status.config(text="処理中…")
            self.update_idletasks()

            # 出力ファイルパス（後から拡張子部分を変える）
            pdf_name = f"{self.pdf_path.stem}{self.pdf_path.suffix}"
            out_file_path = self.out_dir_path / pdf_name

            if self.var_mode.get() == "image":
                # 画像の出力ファイルパスを設定
                out_file_path = out_file_path.with_suffix(".png")
                print(f"out_file_path(img)={out_file_path}")

                # 画像化
                pdf_to_single_image(
                    pdf_path=str(self.pdf_path),
                    page_nums=list(range(page_start, page_end + 1)),
                    out_path=out_file_path,
                    trim_top_px=trim_top,
                    trim_bottom_px=trim_bottom,
                )

            else:
                # Excelに貼り付ける画像ファイルの一時出力先を存在しなければ、作成
                cwd = Path.cwd()
                app_tmp_dir_path = cwd / "app_tmp"
                app_tmp_dir_path.mkdir(parents=True, exist_ok=True)

                # フォルダ内のファイルを削除
                clear_tmp_files(app_tmp_dir_path)

                # Excelに貼り付ける画像ファイルパス
                tmp_img_file_path = app_tmp_dir_path / f"{self.pdf_path.stem}.png"

                # 画像化
                img_paths: list[str] = pdf_to_images(
                    pdf_path=str(self.pdf_path),
                    page_nums=list(range(page_start, page_end + 1)),
                    out_path=tmp_img_file_path,
                    dpi=200,
                    trim_top_px=trim_top,
                    trim_bottom_px=trim_bottom,
                )

                # Excelの出力ファイルパスを設定
                out_file_path = out_file_path.with_suffix(".xlsx")
                print(f"out_file_path(excel)={out_file_path}")

                # Excelに1ページずつ貼り付け
                images_to_excel(img_paths=img_paths, excel_path=out_file_path)

            self.lbl_status.config(text=f"完了: {str(out_file_path)}")
            messagebox.showinfo("完了", f"出力しました。\n{str(out_file_path)}")

        except Exception as e:
            traceback.print_exc()
            self.lbl_status.config(text="失敗")
            messagebox.showerror("エラー", f"失敗しました。\n{e}")

        finally:
            self.btn_run.config(state="normal")


def clear_tmp_files(dir_path: Path):
    for p in dir_path.iterdir():
        if p.is_file():
            try:
                p.unlink()
            except Exception as e:
                print(f"削除失敗: {p} - {e}")
