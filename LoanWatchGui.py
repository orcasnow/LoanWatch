"""LoanWatch GUI launcher.

起動:
    python LoanWatchGui.py
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
MAIN_SCRIPT = Path(__file__).resolve().with_name("LoanWatchMain.py")
EDITOR_SCRIPT = Path(__file__).resolve().with_name("PlayersEditor.py")


def season_choices_for(today):
    current_season_start = today.year if today.month >= 7 else today.year - 1
    max_season_start = current_season_start if today.month <= 6 else current_season_start + 1
    choices = []
    for year in range(current_season_start, max_season_start + 1):
        label = f"{year}/{str(year + 1)[-2:]}シーズン"
        choices.append((label, year))
    return choices


def season_choices():
    return season_choices_for(date.today())


def default_output_name():
    return f"loan_players_{date.today():%Y-%m-%d}"


class LoanWatchGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LoanWatch")
        self.geometry("900x640")
        self.minsize(760, 520)

        self.seasons = season_choices()
        self.season_label_to_value = {label: value for label, value in self.seasons}
        self.log_queue = queue.Queue()
        self.process = None
        self.worker = None

        self.season_var = tk.StringVar(self, value=self.seasons[0][0])
        self.excel_var = tk.BooleanVar(self, value=True)
        self.csv_var = tk.BooleanVar(self, value=True)
        self.out_var = tk.StringVar(self, value=default_output_name())
        self.status_var = tk.StringVar(self, value="待機中")

        self.create_widgets()
        self.after(100, self.poll_log_queue)

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        controls = ttk.LabelFrame(self, text="実行設定", padding=10)
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        controls.columnconfigure(1, weight=1)
        controls.columnconfigure(3, weight=1)

        ttk.Label(controls, text="シーズン").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.season_combo = ttk.Combobox(
            controls,
            textvariable=self.season_var,
            values=[label for label, _value in self.seasons],
            state="readonly",
            width=22,
        )
        self.season_combo.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(controls, text="出力ファイル名").grid(
            row=0, column=2, sticky="w", padx=(18, 6)
        )
        ttk.Entry(controls, textvariable=self.out_var).grid(
            row=0, column=3, sticky="ew", pady=4
        )

        output_frame = ttk.Frame(controls)
        output_frame.grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Label(output_frame, text="出力形式").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Checkbutton(output_frame, text="Excel", variable=self.excel_var).grid(
            row=0, column=1, sticky="w", padx=(0, 10)
        )
        ttk.Checkbutton(output_frame, text="CSV", variable=self.csv_var).grid(
            row=0, column=2, sticky="w"
        )

        actions = ttk.Frame(controls)
        actions.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        actions.columnconfigure(4, weight=1)

        self.run_button = ttk.Button(actions, text="実行", command=self.run_loanwatch_checked)
        self.run_button.grid(row=0, column=0, padx=(0, 8))
        self.stop_button = ttk.Button(
            actions, text="停止", command=self.stop_loanwatch, state="disabled"
        )
        self.stop_button.grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="選手設定を編集", command=self.run_players_editor_checked).grid(
            row=0, column=2, padx=(0, 8)
        )
        ttk.Button(actions, text="ログをクリア", command=self.clear_log).grid(
            row=0, column=3, padx=(0, 8)
        )
        ttk.Label(actions, textvariable=self.status_var).grid(row=0, column=4, sticky="e")

        log_frame = ttk.LabelFrame(self, text="実行ログ", padding=8)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(6, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=20)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

    def build_command(self):
        out_name = self.out_var.get().strip()
        if not out_name:
            raise ValueError("出力ファイル名を入力してください")
        if not self.excel_var.get() and not self.csv_var.get():
            raise ValueError("Excel または CSV を少なくとも1つ選択してください")

        year = self.season_label_to_value[self.season_var.get()]
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--loanwatch-cli", "--year", str(year), "--out", out_name]
        else:
            cmd = [
                sys.executable,
                str(MAIN_SCRIPT),
                "--year",
                str(year),
                "--out",
                out_name,
            ]
        if not self.excel_var.get():
            cmd.append("--no-excel")
        if not self.csv_var.get():
            cmd.append("--no-csv")
        return cmd

    def run_loanwatch(self):
        if self.process is not None:
            return
        try:
            cmd = self.build_command()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            return

        self.clear_log()
        self.append_log("> " + " ".join(cmd))
        self.set_running(True)

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        self.worker = threading.Thread(target=self.run_process_worker, args=(cmd, env), daemon=True)
        self.worker.start()

    def run_process_worker(self, cmd, env):
        try:
            self.process = subprocess.Popen(
                cmd,
                cwd=APP_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.log_queue.put(("log", line.rstrip("\n")))
            return_code = self.process.wait()
            self.log_queue.put(("done", return_code))
        except Exception as e:
            self.log_queue.put(("error", f"{type(e).__name__}: {e}"))

    def stop_loanwatch(self):
        if self.process is None:
            return
        if messagebox.askyesno("停止確認", "実行中の処理を停止しますか？"):
            self.process.terminate()
            self.append_log("停止要求を送信しました")

    def open_players_editor(self):
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--players-editor"]
        else:
            if not EDITOR_SCRIPT.exists():
                messagebox.showerror("起動エラー", f"{EDITOR_SCRIPT} が見つかりません")
                return
            cmd = [sys.executable, str(EDITOR_SCRIPT)]
        try:
            subprocess.Popen(cmd, cwd=APP_DIR)
        except OSError as e:
            messagebox.showerror("起動エラー", str(e))

    def check_players_json(self):
        players_path = APP_DIR / "players.json"
        if players_path.exists():
            return True
        messagebox.showerror(
            "設定ファイルエラー",
            f"{players_path} が見つかりません。\n"
            "LoanWatch.exe と同じフォルダに players.json を置いてください。",
        )
        return False

    def run_players_editor_checked(self):
        if not self.check_players_json():
            return
        self.open_players_editor()

    def run_loanwatch_checked(self):
        if not self.check_players_json():
            return
        self.run_loanwatch()

    def poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "log":
                    self.append_log(payload)
                elif kind == "done":
                    self.append_log(f"終了コード: {payload}")
                    self.set_running(False)
                    self.status_var.set("完了" if payload == 0 else "エラーあり")
                    self.process = None
                elif kind == "error":
                    self.append_log(payload)
                    self.set_running(False)
                    self.status_var.set("起動エラー")
                    self.process = None
        except queue.Empty:
            pass
        self.after(100, self.poll_log_queue)

    def append_log(self, text):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def set_running(self, running):
        self.run_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.status_var.set("実行中" if running else "待機中")


def main():
    app = LoanWatchGui()
    app.mainloop()


def ensure_stdio():
    if sys.stdout is None:
        try:
            sys.stdout = open(1, "w", encoding="utf-8", errors="replace", closefd=False)
        except OSError:
            pass
    elif hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if sys.stderr is None:
        try:
            sys.stderr = open(2, "w", encoding="utf-8", errors="replace", closefd=False)
        except OSError:
            pass
    elif hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def dispatch_internal_mode():
    if len(sys.argv) < 2:
        return False

    mode = sys.argv[1]
    if mode == "--loanwatch-cli":
        ensure_stdio()
        import LoanWatchMain

        sys.argv = ["LoanWatchMain.py", *sys.argv[2:]]
        raise SystemExit(LoanWatchMain.main())

    if mode == "--players-editor":
        import PlayersEditor

        sys.argv = ["PlayersEditor.py", *sys.argv[2:]]
        PlayersEditor.main()
        return True

    return False


if __name__ == "__main__":
    if not dispatch_internal_mode():
        main()
