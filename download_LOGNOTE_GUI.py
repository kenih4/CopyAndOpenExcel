import os
import sys
import json
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import requests


CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "download_LOGNOTE_config.json")


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()


def get_save_dir(kind):
    """保存先フォルダ (configの save_dir の {kind} を置換)"""
    return CONFIG["save_dir"].format(kind=kind)


def get_urls(kind, year, month):
    """ダウンロード候補URL (URL1, URL2) を返す"""
    base_url = CONFIG["base_url"].format(kind=kind, year=year, month=month)
    url1 = f"{base_url}/{year}_{month}.xlsm"
    url2 = f"{base_url}/{year}_{month}_{kind}.xlsm"
    return url1, url2


def download_log(kind, year, month, open_excel_flag):
    """GUIなしでダウンロードを実行する"""
    kind = kind.strip()
    year = year.strip()
    month = month.strip().zfill(2)

    # URL構築
    url1, url2 = get_urls(kind, year, month)

    # 保存先フォルダ
    save_dir = get_save_dir(kind)
    os.makedirs(save_dir, exist_ok=True)

    # 最終的な保存ファイル名
    final_filename = f"{year}_{month}_{kind}.xlsm"
    final_path = os.path.join(save_dir, final_filename)

    print(f"Downloading {kind} {year}/{month}...")
    print(f"Trying URL1: {url1}")

    # URL1 を試す
    success = try_download(url1, final_path)
    source = "URL1"

    if not success:
        print("URL1 not found, trying URL2...")
        success = try_download(url2, final_path)
        source = "URL2"

    if success:
        print(f"Saved: {final_path} ({source})")
        if open_excel_flag:
            open_excel_file(final_path)
        return True
    else:
        print(f"Error: File not found.")
        print(f"URL1: {url1}")
        print(f"URL2: {url2}")
        return False


def try_download(url, save_path):
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            with open(save_path, "wb") as f:
                f.write(response.content)
            return True
    except Exception as e:
        print(f"Download error: {e}")
    return False


def open_excel_file(file_path):
    try:
        os.startfile(file_path)
        print(f"Opened: {file_path}")
    except Exception as e:
        print(f"Warning: Could not open file: {e}")


class LogDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Operation Log Downloader")
        self.root.geometry("500x310")
        self.root.resizable(False, False)

        now = datetime.now()
        current_year = str(now.year)
        current_month = f"{now.month:02d}"

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Kind:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.kind_var = tk.StringVar(value=CONFIG["default_kind"])
        self.kind_combo = ttk.Combobox(
            frame, textvariable=self.kind_var,
            values=CONFIG["kinds"],
            state="readonly", width=15
        )
        self.kind_combo.grid(row=0, column=1, sticky=tk.W, pady=5)

        ttk.Label(frame, text="Year:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.year_var = tk.StringVar(value=current_year)
        ttk.Entry(frame, textvariable=self.year_var, width=18).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(frame, text="Month:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.month_var = tk.StringVar(value=current_month)
        ttk.Entry(frame, textvariable=self.month_var, width=18).grid(row=2, column=1, sticky=tk.W, pady=5)

        ttk.Label(frame, text="Save to:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.save_path_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.save_path_var, foreground="gray").grid(row=3, column=1, sticky=tk.W, pady=5)
        self.update_save_path()

        self.open_excel_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, text="ダウンロード後にExcelで開く",
            variable=self.open_excel_var
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.download_btn = ttk.Button(frame, text="Download", command=self.on_download)
        self.download_btn.grid(row=5, column=0, columnspan=2, pady=10)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.status_var, foreground="blue").grid(row=6, column=0, columnspan=2, pady=5)

        self.kind_var.trace_add("write", lambda *args: self.update_save_path())

    def update_save_path(self):
        self.save_path_var.set(get_save_dir(self.kind_var.get()))

    def on_download(self):
        kind = self.kind_var.get().strip()
        year = self.year_var.get().strip()
        month = self.month_var.get().strip()

        if not kind or not year or not month:
            messagebox.showerror("Error", "Kind, Year, Month をすべて入力してください。")
            return

        month = month.zfill(2)
        url1, url2 = get_urls(kind, year, month)

        save_dir = get_save_dir(kind)
        os.makedirs(save_dir, exist_ok=True)

        final_filename = f"{year}_{month}_{kind}.xlsm"
        final_path = os.path.join(save_dir, final_filename)

        self.status_var.set("Downloading...")
        self.root.update()

        success = try_download(url1, final_path)
        source = "URL1"

        if not success:
            self.status_var.set("URL1 not found, trying URL2...")
            self.root.update()
            success = try_download(url2, final_path)
            source = "URL2"

        if success:
            self.status_var.set(f"Saved: {final_filename} ({source})")
            if self.open_excel_var.get():
                open_excel_file(final_path)
        else:
            self.status_var.set("Download failed.")
            messagebox.showerror("Error", f"ファイルが見つかりませんでした。\n\nURL1: {url1}\nURL2: {url2}")


def main():
    # コマンドライン引数をチェック
    if len(sys.argv) >= 5:
        # GUIなしモード
        kind = sys.argv[1]
        year = sys.argv[2]
        month = sys.argv[3]
        open_flag = sys.argv[4].lower() in ("true", "1", "yes", "on")
        
        success = download_log(kind, year, month, open_flag)
        sys.exit(0 if success else 1)
    else:
        # GUIモード
        root = tk.Tk()
        app = LogDownloaderGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()