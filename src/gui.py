from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pdf import pdf_to_single_image
import fitz


class App(tk.Tk):
    # GUIの構築
    def __init__(self):
        super().__init__()

        # タイトル
        self.title("PDF画像変換")
        # ウィンドウサイズ（横x縦)
        self.geometry("520x300")

        self.pdf_path: Path | None = None
        self.out_path: Path | None = None
        self.page_count = 0

        pad = {"padx": 8, "pady": 6}

        # 行0: PDF選択
        self.btn_pick = ttk.Button(self, text="PDFを選択", command=self.cmd_pick_pdf)
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

        # 行3: 出力先
        self.btn_out = ttk.Button(self, text="保存先を選択", command=self.cmd_pick_out)
        self.btn_out.grid(row=3, column=0, **pad, sticky="w")

        self.lbl_out = ttk.Label(
            self, text="（未選択）", anchor="w", justify="left", wraplength=440
        )
        self.lbl_out.grid(row=3, column=1, columnspan=3, **pad, sticky="ew")

        # 行4: 実行
        self.btn_run = ttk.Button(self, text="実行", command=self.cmd_run)
        self.btn_run.grid(row=4, column=3, **pad, sticky="e")

        # 行5: ステータス
        self.lbl_status = ttk.Label(
            self, text="", anchor="w", justify="left", wraplength=440
        )
        self.lbl_status.grid(row=5, column=0, columnspan=3, **pad, sticky="ew")

        for i in range(4):
            self.grid_columnconfigure(i, weight=0)
        self.grid_columnconfigure(1, weight=1)

    # 「PDFを選択」押下時の処理
    def cmd_pick_pdf(self):
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

        # 選択したPDF名をもとに、保存先のデフォルト値を設定
        default_png = self.pdf_path.with_name(f"{self.pdf_path.stem}.png")
        self.out_path = Path(str(default_png))
        self.lbl_out.config(text=str(default_png))

    # 「保存先を選択」押下時の処理
    def cmd_pick_out(self):
        if not self.pdf_path:
            messagebox.showwarning("警告", "先にPDFを選択してください。")
            return
        path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=".png",
            filetypes=[
                ("PNG 画像", "*.png"),
                ("JPEG 画像", "*.jpg;*.jpeg"),
                ("All Files", "*.*"),
            ],
        )
        if not path:
            return
        self.out_path = Path(path)
        self.lbl_out.config(text=str(self.out_path))

    # 「実行」押下時の処理
    def cmd_run(self):
        if not self.pdf_path:
            messagebox.showwarning("警告", "PDF が選択されていません。")
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

        # PDFの画像化
        try:
            self.btn_run.config(state="disabled")
            self.lbl_status.config(text="処理中…")
            self.update_idletasks()

            pdf_to_single_image(
                pdf_path=str(self.pdf_path),
                page_nums=list(range(page_start, page_end + 1)),
                out_path=str(self.out_path),
                dpi=200,
                trim_top_px=trim_top,
                trim_bottom_px=trim_bottom,
            )

            self.lbl_status.config(text=f"完了: {self.out_path}")
            messagebox.showinfo("完了", f"出力しました。\n{self.out_path}")

        except Exception as e:
            self.lbl_status.config(text="失敗")
            messagebox.showerror("エラー", f"失敗しました。\n{e}")
        finally:
            self.btn_run.config(state="normal")
