import os
import shutil
import subprocess
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD


class PandocConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Pandoc TeX 轉 Word 批次轉換工具")
        self.root.geometry("630x530")
        self.root.resizable(False, False)

        self.selected_files = []
        self.template_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.setup_ui()
        self.setup_drag_and_drop()
        self.check_pandoc_installed()

    def setup_ui(self):
        # 0. 頂部環境需求與下載提示區塊
        frame_notice = tk.Frame(self.root, bg="#FFF8E1", bd=1, relief="solid", padx=10, pady=8)
        frame_notice.pack(fill="x", padx=15, pady=(12, 4))

        row_top = tk.Frame(frame_notice, bg="#FFF8E1")
        row_top.pack(anchor="w")

        tk.Label(
            row_top,
            text="※ 使用程式前須 ",
            bg="#FFF8E1",
            fg="#333333",
            font=("微軟正黑體", 9, "bold"),
        ).pack(side="left")

        lbl_link = tk.Label(
            row_top,
            text="[下載並安裝]",
            bg="#FFF8E1",
            fg="#0066CC",
            cursor="hand2",
            font=("微軟正黑體", 9, "bold", "underline"),
        )
        lbl_link.pack(side="left")
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/jgm/pandoc/releases"))

        tk.Label(
            row_top,
            text=" pandoc-windows-x86_64.msi",
            bg="#FFF8E1",
            fg="#333333",
            font=("微軟正黑體", 9, "bold"),
        ).pack(side="left")

        self.lbl_pandoc_status = tk.Label(
            frame_notice,
            text="正在檢查本機 Pandoc 環境...",
            bg="#FFF8E1",
            fg="gray",
            font=("微軟正黑體", 8),
        )
        self.lbl_pandoc_status.pack(anchor="w", pady=(2, 0))

        # 1. 檔案選擇與拖曳區塊
        frame_files = tk.LabelFrame(
            self.root,
            text="1. 選擇要轉換的 .tex 檔案（支援多選或直接拖曳檔案至此視窗）",
            padx=10,
            pady=10,
        )
        frame_files.pack(fill="x", padx=15, pady=6)

        btn_select = tk.Button(frame_files, text="瀏覽並選擇檔案...", command=self.select_files, width=18)
        btn_select.pack(side="left")

        btn_clear = tk.Button(frame_files, text="清空列表", command=self.clear_files, width=10)
        btn_clear.pack(side="left", padx=(6, 0))

        self.lbl_file_count = tk.Label(frame_files, text="尚未選擇任何檔案（可直接拖曳 .tex 進來）", fg="gray")
        self.lbl_file_count.pack(side="left", padx=12)

        # 2. 進階設定區塊
        frame_opts = tk.LabelFrame(self.root, text="2. 輸出與樣式設定（選填）", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=15, pady=5)

        # 輸出資料夾
        tk.Label(frame_opts, text="輸出資料夾：").grid(row=0, column=0, sticky="w", pady=4)
        tk.Entry(frame_opts, textvariable=self.output_dir, width=43, state="readonly").grid(row=0, column=1, padx=5)
        tk.Button(frame_opts, text="選擇目錄", command=self.select_output_dir).grid(row=0, column=2)
        tk.Label(
            frame_opts,
            text="（留空則預設輸出至原 .tex 檔案相同目錄）",
            fg="gray",
            font=("微軟正黑體", 8),
        ).grid(row=1, column=1, sticky="w")

        # Word 範本檔
        tk.Label(frame_opts, text="Word 樣式範本：").grid(row=2, column=0, sticky="w", pady=4)
        tk.Entry(frame_opts, textvariable=self.template_path, width=43, state="readonly").grid(row=2, column=1, padx=5)
        tk.Button(frame_opts, text="選擇範本", command=self.select_template).grid(row=2, column=2)

        # 3. 執行按鈕
        self.btn_convert = tk.Button(
            self.root,
            text="開始批次轉換為 Word (.docx)",
            command=self.start_conversion_thread,
            bg="#0066CC",
            fg="white",
            font=("微軟正黑體", 11, "bold"),
            height=2,
        )
        self.btn_convert.pack(fill="x", padx=15, pady=10)

        # 4. 執行紀錄視窗
        frame_log = tk.LabelFrame(self.root, text="執行狀態與日誌", padx=10, pady=8)
        frame_log.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        self.log_area = scrolledtext.ScrolledText(frame_log, height=8, state="disabled", font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True)

    def setup_drag_and_drop(self):
        # 註冊整個視窗接受檔案拖曳
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.on_drop_files)

    def check_pandoc_installed(self):
        if shutil.which("pandoc"):
            self.lbl_pandoc_status.config(text="✔ 狀態：本機已偵測到 Pandoc，可正常執行轉換。", fg="#1B7F37")
        else:
            self.lbl_pandoc_status.config(
                text="✘ 狀態：尚未偵測到 Pandoc，請先點擊上方連結下載安裝並重新開啟本程式。",
                fg="#C62828",
            )

    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")

    def add_files(self, file_paths):
        added = 0
        for path in file_paths:
            norm_path = os.path.normpath(path)
            if os.path.isfile(norm_path) and norm_path.lower().endswith(".tex"):
                if norm_path not in self.selected_files:
                    self.selected_files.append(norm_path)
                    added += 1
            else:
                self.log(f"[略過] 非 .tex 檔案：{os.path.basename(norm_path)}")

        if self.selected_files:
            self.lbl_file_count.config(text=f"已載入 {len(self.selected_files)} 個 .tex 檔案", fg="black")
        if added > 0:
            self.log(f"已新增 {added} 個 .tex 檔案（目前共 {len(self.selected_files)} 個）。")

    def on_drop_files(self, event):
        # 使用 tk.splitlist 正確解析含空白字元與大括號的路徑字串
        dropped_paths = self.root.tk.splitlist(event.data)
        self.add_files(dropped_paths)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="選擇 LaTeX 檔案",
            filetypes=[("LaTeX 檔案", "*.tex"), ("所有檔案", "*.*")],
        )
        if files:
            self.add_files(files)

    def clear_files(self):
        self.selected_files.clear()
        self.lbl_file_count.config(text="尚未選擇任何檔案（可直接拖曳 .tex 進來）", fg="gray")
        self.log("已清空待轉換檔案列表。")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="選擇輸出資料夾")
        if directory:
            self.output_dir.set(directory)

    def select_template(self):
        file = filedialog.askopenfilename(
            title="選擇 Word 參考樣式範本",
            filetypes=[("Word 文件", "*.docx")],
        )
        if file:
            self.template_path.set(file)

    def start_conversion_thread(self):
        if not self.selected_files:
            messagebox.showwarning("提示", "請先選擇或拖曳至少一個 .tex 檔案！")
            return

        self.btn_convert.config(state="disabled", text="轉換進行中...")
        threading.Thread(target=self.run_conversion, daemon=True).start()

    def run_conversion(self):
        success_count = 0
        fail_count = 0

        for tex_path in self.selected_files:
            tex_dir = os.path.dirname(tex_path)
            base_name = os.path.splitext(os.path.basename(tex_path))[0]
            target_dir = self.output_dir.get() if self.output_dir.get() else tex_dir
            docx_path = os.path.join(target_dir, f"{base_name}.docx")

            cmd = ["pandoc", tex_path, "-o", docx_path, f"--resource-path=.;{tex_dir}"]

            if self.template_path.get():
                cmd.append(f"--reference-doc={self.template_path.get()}")

            self.log(f"正在轉換: {os.path.basename(tex_path)} ...")

            try:
                result = subprocess.run(
                    cmd,
                    cwd=tex_dir,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    creationflags=0x08000000 if os.name == "nt" else 0,
                )

                if result.returncode == 0:
                    self.log(f"[成功] -> {docx_path}")
                    success_count += 1
                else:
                    self.log(f"[失敗] {os.path.basename(tex_path)} 錯誤訊息:\n{result.stderr.strip()}")
                    fail_count += 1

            except FileNotFoundError:
                self.log("[嚴重錯誤] 找不到 pandoc 指令！請先下載並安裝 pandoc-windows-x86_64.msi。")
                break
            except Exception as e:
                self.log(f"[例外錯誤] {str(e)}")
                fail_count += 1

        self.log("-" * 45)
        self.log(f"批次作業完成！成功：{success_count} 筆，失敗：{fail_count} 筆。\n")
        self.root.after(0, lambda: self.btn_convert.config(state="normal", text="開始批次轉換為 Word (.docx)"))
        self.root.after(0, lambda: messagebox.showinfo("完成", f"轉換完成！\n成功：{success_count} 筆\n失敗：{fail_count} 筆"))


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PandocConverterGUI(root)
    root.mainloop()