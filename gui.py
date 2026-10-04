"""Small, offline Windows interface for the fixed-scope paid service."""
from __future__ import annotations

import json
import os
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from merge_tool import InputError, consolidate, export_result, load_config


class MergeApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=22)
        self.pack(fill="both", expand=True)
        self.paths: list[Path] = []
        self.result = None
        self.config_path = None
        self.busy = False
        self.keys = tk.StringVar()
        self.dedupe = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="파일을 선택하세요. 원본 파일은 변경하지 않습니다.")
        ttk.Label(self, text="엑셀·CSV 파일 취합", font=("맑은 고딕", 21, "bold")).pack(anchor="w")
        ttk.Label(self, text="파일 선택 · 중복 기준 확인 · 검수 · 새 결과 저장", font=("맑은 고딕", 11)).pack(anchor="w", pady=(5, 18))
        actions = ttk.Frame(self)
        actions.pack(fill="x")
        self.add_button = ttk.Button(actions, text="파일 선택", command=self.choose_files)
        self.add_button.pack(side="left")
        self.config_button = ttk.Button(actions, text="설정 파일 불러오기", command=self.choose_config)
        self.config_button.pack(side="left", padx=8)
        self.file_list = tk.Listbox(self, height=6, font=("맑은 고딕", 11), activestyle="none")
        self.file_list.pack(fill="x", pady=12)
        ttk.Label(self, text="중복 검수 기준 열 (선택, 여러 개는 쉼표로 구분)").pack(anchor="w")
        self.key_entry = ttk.Entry(self, textvariable=self.keys, font=("맑은 고딕", 11))
        self.key_entry.pack(fill="x", pady=(5, 8))
        self.check = ttk.Checkbutton(self, text="전체 값이 일치하는 중복 행만 제외하고 기록 남기기", variable=self.dedupe, command=self.invalidate)
        self.check.pack(anchor="w", pady=(0, 12))
        self.keys.trace_add("write", lambda *_: self.invalidate())
        self.preview_button = ttk.Button(self, text="취합 및 검수", command=self.preview)
        self.preview_button.pack(anchor="w")
        ttk.Label(self, textvariable=self.status, wraplength=690, font=("맑은 고딕", 11)).pack(fill="x", pady=16)
        self.output_button = ttk.Button(self, text="새 결과 폴더에 저장", command=self.save, state="disabled")
        self.output_button.pack(anchor="w")
        ttk.Label(self, text="검수대상에는 동일한 키의 값 차이, 필수값 누락, 수식 원문 등이 표시됩니다.\nXLSX 수식은 계산하지 않습니다. CSV 식별자의 선행 0은 보존합니다.", wraplength=690).pack(anchor="w", pady=(18, 0))

    def invalidate(self):
        if self.busy:
            return
        self.result = None
        self.output_button.configure(state="disabled")

    def choose_files(self):
        selected = filedialog.askopenfilenames(title="취합할 파일 선택", filetypes=[("엑셀·CSV", "*.xlsx *.csv *.tsv")])
        if selected:
            self.paths = [Path(p) for p in selected]
            self.file_list.delete(0, "end")
            for path in self.paths:
                self.file_list.insert("end", path.name)
            self.invalidate()
            self.status.set(f"{len(self.paths)}개 파일 선택. 기준을 확인하고 취합 및 검수를 누르세요.")

    def choose_config(self):
        selected = filedialog.askopenfilename(title="고객용 설정 파일", filetypes=[("JSON 설정", "*.json")])
        if selected:
            try:
                config = load_config(selected)
                self.config_path = selected
                self.keys.set(", ".join(config["keys"]))
                self.dedupe.set(config["dedupe_exact"])
                self.invalidate()
                self.status.set(f"설정 불러옴: {Path(selected).name}")
            except (OSError, ValueError) as exc:
                messagebox.showerror("설정 확인", str(exc))

    def set_busy(self, busy):
        self.busy = busy
        for widget in (self.add_button, self.config_button, self.preview_button, self.key_entry, self.check):
            widget.configure(state="disabled" if busy else "normal")
        self.output_button.configure(state="disabled")

    def preview(self):
        if not self.paths:
            messagebox.showwarning("파일 선택", "먼저 파일을 선택하세요.")
            return
        try:
            config = load_config(self.config_path)
            config["keys"] = [c.strip() for c in self.keys.get().split(",") if c.strip()]
            config["dedupe_exact"] = self.dedupe.get()
        except (OSError, ValueError) as exc:
            messagebox.showerror("설정 확인", str(exc))
            return
        self.result = None
        self.set_busy(True)
        self.status.set("로컬에서 처리 중입니다.")
        paths = list(self.paths)
        def work():
            try:
                result = consolidate(paths, config)
                self.after(0, lambda: self.finish(result, None))
            except Exception as exc:
                error_text = str(exc)
                self.after(0, lambda: self.finish(None, error_text))
        threading.Thread(target=work, daemon=True).start()

    def finish(self, result, error):
        self.set_busy(False)
        if error:
            self.status.set(f"처리 중단: {error}")
            return
        self.result = result
        report = result.report
        self.status.set(f"입력 {report['input_rows']:,}행 → 결과 {report['output_rows']:,}행\n제외 {report['excluded_rows']:,}행 · 검수 {report['review_rows']:,}행 · 동일키 값 차이 {report['conflicting_key_groups']:,}그룹")
        self.output_button.configure(state="normal")

    def save(self):
        if self.result is None:
            return
        parent = filedialog.askdirectory(title="새 결과 폴더를 만들 위치")
        if not parent:
            return
        destination = Path(parent) / f"취합결과_{datetime.now():%Y%m%d_%H%M%S_%f}"
        try:
            output = export_result(self.result, destination)
            self.status.set(f"저장 완료: {output}\n검수대상 시트를 확인하세요.")
            messagebox.showinfo("저장 완료", f"{output}\n\n검수대상 시트와 제외기록을 확인하세요.")
        except Exception as exc:
            messagebox.showerror("저장 중단", str(exc))


def run_gui():
    root = tk.Tk()
    root.title("파일 취합 도구 0.1")
    root.geometry("780x630")
    root.minsize(660, 560)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TButton", font=("맑은 고딕", 11), padding=(14, 9))
    style.configure("TLabel", font=("맑은 고딕", 10))
    style.configure("TCheckbutton", font=("맑은 고딕", 10))
    MergeApp(root)
    root.mainloop()
