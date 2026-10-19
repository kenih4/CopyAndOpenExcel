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
    "bg": "#1e1e1e", "card": "#262628", "fg": "#e6e6e6", "muted": "#9a9a9a",
    "field": "#333336", "border": "#3d3d40",
    "button": "#38383c", "button_active": "#48484d", "accent": "#2f6db5", "accent_active": "#3b80d0",
    "link": "#6cb6ff", "disabled": "#777777",
}

GREP_SCRIPT = "c:/Users/kenic/Dropbox/gitdir/operation_log_NEW/excelgrep_by_XMLparse.sh"
ICON_NAME = "CopyAndOpenExcel.ico"
VSCODE_EXE = r"C:\Users\kenic\AppData\Local\Programs\Microsoft VS Code\Code.exe"
SELF_DIR = r"C:\Users\kenic\Dropbox\gitdir\CopyAndOpenExcel"
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
    try:
        success = try_download(url1, final_path)
        source = "URL1"

        if not success:
            print("URL1 not found, trying URL2...")
            success = try_download(url2, final_path)
            source = "URL2"
    except FileLockedError as e:
        print(f"Error: {e}")
        return False

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


class FileLockedError(Exception):
    """保存先ファイルが他のアプリ (Excel等) で開かれていて上書きできない"""

    def __init__(self, path):
        super().__init__(
            f"保存先のファイルが Excel などで開かれているため上書きできません。\n"
            f"ファイルを閉じてからもう一度実行してください。\n\n{path}"
        )


def try_download(url, save_path, progress=None):
    """ダウンロードして save_path に保存。progress(受信バイト数, 全体バイト数 or None) を随時呼ぶ"""
    tmp_path = save_path + ".part"
    # Excel等で開かれていて上書きできないファイルは、ダウンロード前に検出する
    if os.path.exists(save_path):
        try:
            open(save_path, "ab").close()
        except OSError as e:
            raise FileLockedError(save_path) from e
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
        try:
            os.replace(tmp_path, save_path)
        except PermissionError as e:
            os.remove(tmp_path)
            raise FileLockedError(save_path) from e
        return True
    except FileLockedError:
        raise
    except Exception as e:
        print(f"Download error: {e}")
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except OSError:
        pass
    return False


def open_excel_file(file_path):
    """関連付けでExcelを開き、開いたブックを別スレッドで最大化する (最大化に失敗しても開くのは妨げない)"""
    try:
        os.startfile(file_path)
        print(f"Opened: {file_path}")
    except Exception as e:
        print(f"Warning: Could not open file: {e}")
        return
    threading.Thread(target=maximize_excel_window, args=(file_path,), daemon=True).start()


def maximize_excel_window(file_path):
    """起動済みExcelに接続し、指定ブックのウィンドウを最大化する。開くまで最大20秒待つ"""
    ps_script = (
        "$path=$env:LOGNOTE_FILE;"
        "for($i=0;$i -lt 40;$i++){"
        " try{"
        "  $xl=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');"
        "  foreach($wb in $xl.Workbooks){"
        "   if($wb.FullName -ieq $path){"
        "    $xl.WindowState=-4137;"  # xlMaximized (アプリ)
        "    $wb.Windows.Item(1).WindowState=-4137;"  # ブックのウィンドウも最大化
        "    exit 0"
        "   }"
        "  }"
        " }catch{}"
        " Start-Sleep -Milliseconds 500"
        "};"
        "exit 1"
    )
    try:
        env = dict(os.environ, LOGNOTE_FILE=os.path.abspath(file_path))
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            env=env, capture_output=True, timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except Exception as e:
        print(f"Warning: could not maximize Excel window: {e}")


class LogDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Operation Log Downloader")
        self.setup_icon()
        self.root.resizable(False, False)
        self.setup_dark_theme()

        now = datetime.now()
        self.kind_var = tk.StringVar(value=CONFIG["default_kind"])
        self.year_var = tk.StringVar(value=str(now.year))
        self.month_var = tk.StringVar(value=f"{now.month:02d}")
        self.save_path_var = tk.StringVar()
        self.open_excel_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")

        outer = ttk.Frame(self.root, padding=(16, 14, 16, 0))
        outer.pack(fill=tk.BOTH, expand=True)
        tk.Frame(outer, width=620, height=0, bg=DARK["bg"]).pack()  # ウィンドウの最小幅を確保する

        # ---- 対象 (Kind / Year / Month / 保存先) ----
        target = self.make_card(outer, "TARGET")

        kind_row = ttk.Frame(target, style="Card.TFrame")
        kind_row.pack(fill=tk.X, pady=(0, 10))
        # indicatoron=0 でボタン表示にし、選択中は押し込み状態 (アクセント色) になる
        for i, kind in enumerate(CONFIG["kinds"]):
            kind_row.columnconfigure(i, weight=1, uniform="kind")  # 幅を均等にする
            tk.Radiobutton(
                kind_row, text=kind, value=kind, variable=self.kind_var,
                indicatoron=0, pady=6, relief=tk.FLAT, bd=0, font=("", 10, "bold"),
                bg=DARK["button"], fg=DARK["fg"], selectcolor=DARK["accent"],
                activebackground=DARK["button_active"], activeforeground=DARK["fg"],
                highlightthickness=0, cursor="hand2"
            ).grid(row=0, column=i, sticky="ew", padx=2)

        date_row = ttk.Frame(target, style="Card.TFrame")
        date_row.pack(fill=tk.X, pady=(0, 10))
        self.make_stepper(date_row, "Year", self.year_var, 12).pack(side=tk.LEFT, expand=True)
        self.make_stepper(date_row, "Month", self.month_var, 1).pack(side=tk.LEFT, expand=True)

        save_row = ttk.Frame(target, style="Card.TFrame")
        save_row.pack(fill=tk.X)
        ttk.Label(save_row, text="Save to", style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        save_label = ttk.Label(save_row, textvariable=self.save_path_var, style="Link.TLabel", cursor="hand2")
        save_label.pack(side=tk.LEFT)
        save_label.bind("<Button-1>", lambda e: self.open_save_dir())
        self.update_save_path()

        # ---- 実行 ----
        actions = self.make_card(outer, "ACTIONS")
        grid = ttk.Frame(actions, style="Card.TFrame")
        grid.pack(fill=tk.X)
        grid.columnconfigure((0, 1), weight=1, uniform="act")
        self.open_btn = ttk.Button(grid, text="Open  (ローカル優先)", style="Primary.TButton",
                                   command=lambda: self.on_download(force=False))
        self.download_btn = ttk.Button(grid, text="Re-download  (再取得)",
                                       command=lambda: self.on_download(force=True))
        self.log_btn = ttk.Button(grid, text="Open + Log出力  (ローカル優先)",
                                  command=lambda: self.on_download(force=False, log=True))
        self.log_dl_btn = ttk.Button(grid, text="Re-download + Log出力",
                                     command=lambda: self.on_download(force=True, log=True))
        self.open_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=(0, 8))
        self.download_btn.grid(row=0, column=1, sticky="ew", padx=(4, 0), pady=(0, 8))
        self.log_btn.grid(row=1, column=0, sticky="ew", padx=(0, 4))
        self.log_dl_btn.grid(row=1, column=1, sticky="ew", padx=(4, 0))
        ttk.Checkbutton(actions, text="ダウンロード後にExcelで開く", variable=self.open_excel_var,
                        style="Card.TCheckbutton").pack(anchor=tk.W, pady=(10, 0))

        # ---- ツール ----
        tools = self.make_card(outer, "TOOLS")
        self.make_button_row(tools, [
            ("operation_log_NEW", self.open_operation_log_in_vscode),
            ("ログ検索", self.open_logsearch),
        ])

        # ---- 設定 ----
        settings = self.make_card(outer, "SETTINGS")
        self.make_button_row(settings, [
            ("vscode CopyAndOpenExcel", self.open_self_in_vscode),
        ])

        # ---- ステータスバー ----
        status_bar = tk.Frame(self.root, bg=DARK["card"])
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(12, 0))
        ttk.Label(status_bar, textvariable=self.status_var, style="Status.TLabel").pack(
            side=tk.LEFT, padx=16, pady=6)

        self.kind_var.trace_add("write", lambda *args: self.update_save_path())

    def make_card(self, parent, title):
        """見出し付きのカード (セクション) を作り、中身を入れるフレームを返す"""
        wrapper = tk.Frame(parent, bg=DARK["card"], highlightthickness=1,
                           highlightbackground=DARK["border"], highlightcolor=DARK["border"])
        wrapper.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(wrapper, text=title, style="Section.TLabel").pack(anchor=tk.W, padx=14, pady=(10, 6))
        body = ttk.Frame(wrapper, style="Card.TFrame", padding=(14, 0, 14, 14))
        body.pack(fill=tk.X)
        return body

    def make_button_row(self, parent, buttons, columns=2):
        """同じ幅のボタンを横に並べる。ボタンが少なくても列数分の幅で揃える"""
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X)
        row.columnconfigure(tuple(range(columns)), weight=1, uniform="btn")
        for i, (text, cmd) in enumerate(buttons):
            ttk.Button(row, text=text, command=cmd).grid(
                row=0, column=i, sticky="ew",
                padx=(0 if i == 0 else 4, 0 if i == columns - 1 else 4))

    def make_stepper(self, parent, label, var, step):
        """ラベル + [-] [入力欄] [+] のセット。step=12 で年、1 で月"""
        box = ttk.Frame(parent, style="Card.TFrame")
        ttk.Label(box, text=label, style="Muted.TLabel", width=6).pack(side=tk.LEFT)
        ttk.Button(box, text="−", width=3, command=lambda: self.change_month(-step)).pack(side=tk.LEFT)
        ttk.Entry(box, textvariable=var, width=7, justify="center", font=("", 11, "bold")).pack(side=tk.LEFT, padx=4, ipady=3)
        ttk.Button(box, text="+", width=3, command=lambda: self.change_month(step)).pack(side=tk.LEFT)
        return box

    def open_self_in_vscode(self):
        try:
            subprocess.Popen([VSCODE_EXE, SELF_DIR])
        except Exception as e:
            messagebox.showerror("Error", f"VS Code を起動できませんでした。\n{e}")

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

    def setup_icon(self):
        """スクリプトと同じフォルダの CopyAndOpenExcel.ico があればアイコンに使う"""
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ICON_NAME)
        if not os.path.isfile(icon_path):
            return
        try:
            # python.exe とは別のアプリとして扱わせ、タスクバーに専用アイコンを出す
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("CopyAndOpenExcel.LogDownloader")
        except Exception:
            pass
        try:
            self.root.iconbitmap(default=icon_path)
        except tk.TclError as e:
            print(f"Warning: could not load icon: {e}")

    def setup_dark_theme(self):
        self.root.configure(bg=DARK["bg"])
        style = ttk.Style()
        style.theme_use("clam")  # 色指定が効くテーマ
        style.configure(".", background=DARK["bg"], foreground=DARK["fg"],
                        fieldbackground=DARK["field"], bordercolor=DARK["border"],
                        lightcolor=DARK["button"], darkcolor=DARK["button"],
                        troughcolor=DARK["field"], insertcolor=DARK["fg"])
        style.configure("Card.TFrame", background=DARK["card"])
        style.configure("Section.TLabel", background=DARK["card"], foreground=DARK["muted"], font=("", 8, "bold"))
        style.configure("Muted.TLabel", background=DARK["card"], foreground=DARK["muted"])
        style.configure("Link.TLabel", background=DARK["card"], foreground=DARK["link"])
        style.configure("Status.TLabel", background=DARK["card"], foreground=DARK["link"])
        style.configure("TButton", background=DARK["button"], foreground=DARK["fg"], padding=(10, 6),
                        borderwidth=0, focuscolor=DARK["button"])
        style.map("TButton",
                  background=[("disabled", DARK["bg"]), ("pressed", DARK["accent"]), ("active", DARK["button_active"])],
                  foreground=[("disabled", DARK["disabled"])])
        style.configure("Primary.TButton", background=DARK["accent"], foreground="#ffffff",
                        font=("", 9, "bold"), focuscolor=DARK["accent"])
        style.map("Primary.TButton",
                  background=[("disabled", DARK["bg"]), ("pressed", DARK["accent"]), ("active", DARK["accent_active"])],
                  foreground=[("disabled", DARK["disabled"])])
        style.configure("TEntry", fieldbackground=DARK["field"], foreground=DARK["fg"], bordercolor=DARK["border"])
        style.configure("Card.TCheckbutton", background=DARK["card"], foreground=DARK["fg"],
                        indicatorbackground=DARK["field"], indicatorforeground=DARK["fg"])
        style.map("Card.TCheckbutton",
                  background=[("active", DARK["card"])],
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