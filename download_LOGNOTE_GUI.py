import os
import sys
import json
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import requests


CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "download_LOGNOTE_config.json")


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()

DARK = {
    "bg": "#1e1e1e", "fg": "#e0e0e0", "field": "#2d2d2d", "border": "#555555",
    "button": "#3a3a3a", "button_active": "#4a4a4a", "accent": "#2f6db5",
    "link": "#6cb6ff", "disabled": "#777777",
}

GREP_SCRIPT = "c:/Users/kenic/Dropbox/gitdir/operation_log_NEW/excelgrep_by_XMLparse.sh"
VSCODE_EXE = r"C:\Users\kenic\AppData\Local\Programs\Microsoft VS Code\Code.exe"
OPERATION_LOG_DIR = r"C:\Users\kenic\Dropbox\gitdir\operation_log_NEW"
LOGSEARCH_URL = "http://saclaopr19.spring8.or.jp/~logsearch/"
GIT_BASH = r"C:\Program Files\Git\bin\bash.exe"


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


def try_download(url, save_path, progress=None):
    """ダウンロードして save_path に保存。progress(受信バイト数, 全体バイト数 or None) を随時呼ぶ"""
    tmp_path = save_path + ".part"
    try:
        with requests.get(url, timeout=30, stream=True) as response:
            if response.status_code != 200:
                return False
            total = int(response.headers.get("Content-Length") or 0) or None
            received = 0
            with open(tmp_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=256 * 1024):
                    f.write(chunk)
                    received += len(chunk)
                    if progress:
                        progress(received, total)
        # 完了してから置き換える (失敗時に既存ファイルを壊さない)
        os.replace(tmp_path, save_path)
        return True
    except Exception as e:
        print(f"Download error: {e}")
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except OSError:
        pass
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
        self.root.geometry("720x400")
        self.root.resizable(False, False)

        now = datetime.now()
        current_year = str(now.year)
        current_month = f"{now.month:02d}"

        frame = ttk.Frame(root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        self.setup_dark_theme()

        ttk.Label(frame, text="Kind:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.kind_var = tk.StringVar(value=CONFIG["default_kind"])
        kind_frame = ttk.Frame(frame)
        kind_frame.grid(row=0, column=1, sticky=tk.W, pady=5)
        # indicatoron=0 でボタン表示にし、選択中は押し込み状態になる
        for kind in CONFIG["kinds"]:
            tk.Radiobutton(
                kind_frame, text=kind, value=kind, variable=self.kind_var,
                indicatoron=0, width=8, pady=3, relief=tk.RAISED, bd=2,
                bg=DARK["button"], fg=DARK["fg"], selectcolor=DARK["accent"],
                activebackground=DARK["button_active"], activeforeground=DARK["fg"],
                highlightthickness=0
            ).pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="Year:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.year_var = tk.StringVar(value=current_year)
        year_frame = ttk.Frame(frame)
        year_frame.grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Button(year_frame, text="-", width=3, command=lambda: self.change_month(-12)).pack(side=tk.LEFT)
        ttk.Entry(year_frame, textvariable=self.year_var, width=10, justify="center").pack(side=tk.LEFT, padx=3)
        ttk.Button(year_frame, text="+", width=3, command=lambda: self.change_month(12)).pack(side=tk.LEFT)

        ttk.Label(frame, text="Month:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.month_var = tk.StringVar(value=current_month)
        month_frame = ttk.Frame(frame)
        month_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Button(month_frame, text="-", width=3, command=lambda: self.change_month(-1)).pack(side=tk.LEFT)
        ttk.Entry(month_frame, textvariable=self.month_var, width=10, justify="center").pack(side=tk.LEFT, padx=3)
        ttk.Button(month_frame, text="+", width=3, command=lambda: self.change_month(1)).pack(side=tk.LEFT)

        ttk.Label(frame, text="Save to:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.save_path_var = tk.StringVar()
        save_label = ttk.Label(frame, textvariable=self.save_path_var, foreground=DARK["link"], cursor="hand2")
        save_label.grid(row=3, column=1, sticky=tk.W, pady=5)
        save_label.bind("<Button-1>", lambda e: self.open_save_dir())
        self.update_save_path()

        self.open_excel_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame, text="ダウンロード後にExcelで開く",
            variable=self.open_excel_var
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=3, pady=10)
        self.open_btn = ttk.Button(btn_frame, text="Open (ローカル優先)", command=lambda: self.on_download(force=False))
        self.open_btn.pack(side=tk.LEFT, padx=5)
        self.download_btn = ttk.Button(btn_frame, text="Re-download (再ダウンロード)", command=lambda: self.on_download(force=True))
        self.download_btn.pack(side=tk.LEFT, padx=5)

        btn_frame2 = ttk.Frame(frame)
        btn_frame2.grid(row=6, column=0, columnspan=3, pady=(0, 10))
        self.log_btn = ttk.Button(btn_frame2, text="Open + Log出力 (ローカル優先)", command=lambda: self.on_download(force=False, log=True))
        self.log_btn.pack(side=tk.LEFT, padx=5)
        self.log_dl_btn = ttk.Button(btn_frame2, text="Re-download + Log出力", command=lambda: self.on_download(force=True, log=True))
        self.log_dl_btn.pack(side=tk.LEFT, padx=5)

        tool_frame = ttk.Frame(frame)
        tool_frame.grid(row=1, column=2, rowspan=2, sticky=tk.N, padx=(25, 0))
        ttk.Button(tool_frame, text="operation_log_NEW", width=18, command=self.open_operation_log_in_vscode).pack(pady=(0, 6))
        ttk.Button(tool_frame, text="ログ検索", width=18, command=self.open_logsearch).pack()

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.status_var, foreground=DARK["link"]).grid(row=7, column=0, columnspan=3, pady=5)

        self.kind_var.trace_add("write", lambda *args: self.update_save_path())

    def open_operation_log_in_vscode(self):
        try:
            subprocess.Popen([VSCODE_EXE, OPERATION_LOG_DIR])
        except Exception as e:
            messagebox.showerror("Error", f"VS Code を起動できませんでした。\n{e}")

    def open_logsearch(self):
        try:
            # 既定ブラウザに関わらず Edge で開く
            subprocess.Popen(["cmd", "/c", "start", "", "msedge", LOGSEARCH_URL],
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception as e:
            messagebox.showerror("Error", f"Edge を起動できませんでした。\n{e}")

    def setup_dark_theme(self):
        self.root.configure(bg=DARK["bg"])
        style = ttk.Style()
        style.theme_use("clam")  # 色指定が効くテーマ
        style.configure(".", background=DARK["bg"], foreground=DARK["fg"],
                        fieldbackground=DARK["field"], bordercolor=DARK["border"],
                        lightcolor=DARK["button"], darkcolor=DARK["button"],
                        troughcolor=DARK["field"], insertcolor=DARK["fg"])
        style.configure("TButton", background=DARK["button"], foreground=DARK["fg"], padding=4)
        style.map("TButton",
                  background=[("disabled", DARK["bg"]), ("pressed", DARK["accent"]), ("active", DARK["button_active"])],
                  foreground=[("disabled", DARK["disabled"])])
        style.configure("TEntry", fieldbackground=DARK["field"], foreground=DARK["fg"])
        style.configure("TCheckbutton", background=DARK["bg"], foreground=DARK["fg"],
                        indicatorbackground=DARK["field"], indicatorforeground=DARK["fg"])
        style.map("TCheckbutton",
                  background=[("active", DARK["bg"])],
                  indicatorcolor=[("selected", DARK["accent"]), ("!selected", DARK["field"])])

    def change_month(self, delta):
        """月単位で増減 (年をまたぐ場合は年も更新)。delta=±12 で年のみ変更"""
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            return
        index = year * 12 + (month - 1) + delta
        self.year_var.set(str(index // 12))
        self.month_var.set(f"{index % 12 + 1:02d}")

    def open_save_dir(self):
        """保存先フォルダをエクスプローラーで開く (なければ作成)"""
        save_dir = get_save_dir(self.kind_var.get())
        try:
            os.makedirs(save_dir, exist_ok=True)
            os.startfile(save_dir)
        except Exception as e:
            messagebox.showerror("Error", f"フォルダを開けませんでした。\n{e}")

    def update_save_path(self):
        self.save_path_var.set(get_save_dir(self.kind_var.get()))

    def set_status(self, text):
        # ワーカースレッドからも安全に呼べるようメインスレッドに依頼する
        self.root.after(0, lambda: self.status_var.set(text))

    def set_buttons_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for btn in (self.open_btn, self.download_btn, self.log_btn, self.log_dl_btn):
            btn.config(state=state)

    def finish(self, path, open_flag, log):
        """取得済みのファイルに対する後処理 (ログ出力 / Excelで開く)"""
        if log:
            self.run_log_script(path)
        elif open_flag:
            open_excel_file(path)

    def run_log_script(self, path):
        """新しいコンソールウィンドウでスクリプトを実行 (Yes/No の対話入力ができる)"""
        self.set_status("Log出力中 (ターミナルで操作してください)...")
        # 終了後に結果を読めるよう、Enterを押すまでウィンドウを閉じない
        wrapper = 'bash "$0" "$@"; rc=$?; echo; read -r -p "終了しました。Enterで閉じます" _; exit $rc'
        result = subprocess.run(
            [GIT_BASH, "-c", wrapper, GREP_SCRIPT, "-l", path.replace("\\", "/")],
            cwd=os.path.dirname(GREP_SCRIPT),
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )
        if result.returncode == 0:
            self.set_status(f"Log出力完了: {os.path.basename(path)}")
        else:
            self.set_status(f"Log出力失敗 (終了コード {result.returncode})")

    def on_download(self, force, log=False):
        kind = self.kind_var.get().strip()
        year = self.year_var.get().strip()
        month = self.month_var.get().strip()

        if not kind or not year or not month:
            messagebox.showerror("Error", "Kind, Year, Month をすべて入力してください。")
            return

        month = month.zfill(2)
        open_flag = self.open_excel_var.get()
        self.set_buttons_enabled(False)
        threading.Thread(
            target=self.worker, args=(kind, year, month, force, open_flag, log), daemon=True
        ).start()

    def worker(self, kind, year, month, force, open_flag, log):
        url1, url2 = get_urls(kind, year, month)
        save_dir = get_save_dir(kind)
        final_filename = f"{year}_{month}_{kind}.xlsm"
        final_path = os.path.join(save_dir, final_filename)

        def progress(received, total):
            mb = received / 1024 / 1024
            if total:
                self.set_status(f"Downloading... {mb:.1f} / {total / 1024 / 1024:.1f} MB ({received * 100 // total}%)")
            else:
                self.set_status(f"Downloading... {mb:.1f} MB")

        try:
            os.makedirs(save_dir, exist_ok=True)

            if not force and os.path.isfile(final_path):
                self.set_status(f"Local file: {final_filename}")
                self.finish(final_path, open_flag, log)
                return

            self.set_status("Downloading...")
            success = try_download(url1, final_path, progress)
            source = "URL1"
            if not success:
                self.set_status("URL1 not found, trying URL2...")
                success = try_download(url2, final_path, progress)
                source = "URL2"

            if success:
                self.set_status(f"Saved: {final_filename} ({source})")
                self.finish(final_path, open_flag, log)
            else:
                self.set_status("Download failed.")
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"ファイルが見つかりませんでした。\n\nURL1: {url1}\nURL2: {url2}"))
        except Exception as e:
            self.set_status("Error.")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.set_buttons_enabled(True))


def hide_own_console():
    """ダブルクリック起動で自動的に開いたコンソールを閉じる (ターミナルから起動した場合は何もしない)"""
    if os.name != "nt":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        pids = (ctypes.c_uint * 2)()
        # コンソールに接続しているのが自プロセスだけなら、自分専用のコンソール
        if kernel32.GetConsoleProcessList(pids, 2) == 1:
            hwnd = kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass


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
        hide_own_console()
        root = tk.Tk()
        app = LogDownloaderGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()