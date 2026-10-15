# -*- coding: utf-8 -*-
"""원시자료 정리봇 화면 (tkinter)

A. txt 폴더 → _01, _02, _03
   사용자가 _03 파일 EVENT 시트에 PT_한국어/SOC_한국어 입력
B. _03 파일 → _10, _20, _30
   사용자가 _30 파일 확인 후 LLT_한국어/반영여부 직접 입력

명령줄 실행(검증용):
  python app.py --step A --txt "폴더" --product 제품명 --drug-name 의약품명 [--out 폴더] [--date YYYYMMDD]
  python app.py --step B --xlsx "_03 파일" --sheet-prefix 시트명 [--drug-name 의약품명] [--out 폴더] [--date YYYYMMDD]
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import queue
import re
import sys
import threading
import time
from pathlib import Path

import engine
from engine import VERSION, BotError

LOG_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "원시자료정리봇" / "logs"


def setup_logging() -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / ("rawbot_" + time.strftime("%Y%m%d_%H%M%S") + ".log")
    logging.basicConfig(filename=log_path, encoding="utf-8", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    logging.info("%s 시작", VERSION)
    return log_path


def explain(exc: BaseException) -> str:
    """사용자에게 보여줄 오류 설명"""
    if isinstance(exc, BotError):
        return str(exc)
    if isinstance(exc, PermissionError):
        return ("파일을 저장하거나 읽을 수 없습니다.\n"
                "같은 이름의 엑셀 파일이 열려 있으면 닫고 다시 실행하세요.\n\n" + str(exc))
    return f"예상하지 못한 오류가 발생했습니다.\n{type(exc).__name__}: {str(exc)[:400]}"


def open_path(path: Path) -> None:
    if hasattr(os, "startfile"):
        os.startfile(path)


# ════════════════════════════════════════════════════════════════════
# 화면
# ════════════════════════════════════════════════════════════════════
class QueueHandler(logging.Handler):
    def __init__(self, events: queue.Queue):
        super().__init__(level=logging.INFO)
        self.events = events

    def emit(self, record):
        self.events.put(("log", (record.levelno, record.getMessage())))


class App:
    def __init__(self, root, log_path: Path):
        import tkinter as tk
        from tkinter import ttk
        self.tk, self.ttk = tk, ttk
        self.root, self.log_path = root, log_path
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.last_out: Path | None = None
        logging.getLogger().addHandler(QueueHandler(self.events))

        root.title(VERSION)
        root.geometry("900x720")
        root.minsize(760, 560)
        today = dt.date.today().strftime("%Y%m%d")

        # 상단 공통 버튼
        bar = ttk.Frame(root, padding=(8, 8, 8, 0)); bar.pack(fill="x")
        ttk.Label(bar, text=VERSION, font=("맑은 고딕", 12, "bold")).pack(side="left")
        ttk.Button(bar, text="로그 폴더", command=lambda: open_path(LOG_DIR)).pack(side="right", padx=3)
        ttk.Button(bar, text="로그 열기", command=lambda: open_path(self.log_path)).pack(side="right", padx=3)
        self.open_out_button = ttk.Button(bar, text="결과 폴더 열기", state="disabled",
                                          command=lambda: self.last_out and open_path(self.last_out))
        self.open_out_button.pack(side="right", padx=3)

        nb = ttk.Notebook(root, padding=8); nb.pack(fill="x")
        self.actions = []

        # ── A 탭 ─────────────────────────────────────────────
        a = ttk.Frame(nb, padding=10); nb.add(a, text="  A. 01~03 만들기  ")
        self.a_txt = tk.StringVar(); self.a_out = tk.StringVar()
        self.a_product = tk.StringVar(); self.a_drug = tk.StringVar(); self.a_date = tk.StringVar(value=today)
        self.path_row(a, 0, "txt 폴더 (13개)", self.a_txt, self.pick_txt_dir)
        self.entry_row(a, 1, "제품명 (파일 이름용)", self.a_product, "예: 그리타존정15mg")
        self.entry_row(a, 2, "의약품명 (DRUG 시트 표시)", self.a_drug, "예: 그리타존정15밀리그램  (비우면 제품명)")
        self.entry_row(a, 3, "파일 날짜", self.a_date, "YYYYMMDD")
        self.path_row(a, 4, "결과 저장 폴더", self.a_out, lambda: self.pick_dir(self.a_out))
        ttk.Label(a, foreground="#555555", justify="left", text=(
            "실행 후 _03 파일 EVENT 시트에 PT_한국어, SOC_한국어를 입력·저장하고 B 탭으로 진행합니다.\n"
            "LLT_한국어, 반영여부는 B 실행 후 _30 파일을 확인하며 직접 입력합니다.")
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.action_button(a, "A 실행  (txt → 01, 02, 03)", self.run_a).grid(row=6, column=0, columnspan=3, sticky="e", pady=(8, 0))
        a.columnconfigure(1, weight=1)

        # ── B 탭 ─────────────────────────────────────────────
        b = ttk.Frame(nb, padding=10); nb.add(b, text="  B. 10~30 만들기  ")
        self.b_xlsx = tk.StringVar(); self.b_out = tk.StringVar()
        self.b_product = tk.StringVar(); self.b_prefix = tk.StringVar(); self.b_date = tk.StringVar(value=today)
        self.b_drug = tk.StringVar()
        self.path_row(b, 0, "_03 파일 (입력 완료)", self.b_xlsx, self.pick_03)
        self.entry_row(b, 1, "제품명 (파일 이름용)", self.b_product, "_03 파일 이름에서 자동 입력")
        ttk.Label(b, text="의심/병용 나눌 의약품명").grid(row=2, column=0, sticky="w", pady=3)
        self.b_drug_box = ttk.Combobox(b, textvariable=self.b_drug, state="readonly")
        self.b_drug_box.grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Label(b, text="DRUG 시트에서 자동", foreground="#777777").grid(row=2, column=2, sticky="w", padx=6)
        self.entry_row(b, 3, "시트 이름 앞부분", self.b_prefix, "예: 그리타존정  →  그리타존정_의심 / _병용")
        self.entry_row(b, 4, "파일 날짜", self.b_date, "YYYYMMDD")
        self.path_row(b, 5, "결과 저장 폴더", self.b_out, lambda: self.pick_dir(self.b_out))
        self.action_button(b, "B 실행  (03 → 10, 20, 30)", self.run_b).grid(row=6, column=0, columnspan=3, sticky="e", pady=(8, 0))
        b.columnconfigure(1, weight=1)

        # ── 진행 기록 ───────────────────────────────────────
        logf = ttk.LabelFrame(root, text="진행 기록", padding=6); logf.pack(fill="both", expand=True, padx=8)
        self.text = tk.Text(logf, wrap="none", height=14, state="disabled", font=("맑은 고딕", 9))
        self.text.tag_configure("warn", foreground="#B35C00")
        self.text.tag_configure("error", foreground="#C00000")
        self.text.tag_configure("done", foreground="#1E6B2E")
        sy = ttk.Scrollbar(logf, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y"); self.text.pack(side="left", fill="both", expand=True)

        self.progress = ttk.Progressbar(root, mode="indeterminate"); self.progress.pack(fill="x", padx=8, pady=(6, 0))
        self.status = tk.StringVar(value="A 탭에서 txt 폴더를 선택하세요.")
        ttk.Label(root, textvariable=self.status, padding=8).pack(fill="x")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)

    # ── 화면 구성 도우미 ───────────────────────────────────
    def entry_row(self, frame, row, label, var, hint):
        ttk = self.ttk
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
        e = ttk.Entry(frame, textvariable=var); e.grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Label(frame, text=hint, foreground="#777777").grid(row=row, column=2, sticky="w", padx=6)
        self.actions.append(e)

    def path_row(self, frame, row, label, var, command):
        ttk = self.ttk
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
        e = ttk.Entry(frame, textvariable=var); e.grid(row=row, column=1, sticky="ew", pady=3)
        b = ttk.Button(frame, text="찾아보기…", command=command); b.grid(row=row, column=2, sticky="w", padx=6)
        self.actions += [e, b]

    def action_button(self, frame, text, command):
        b = self.ttk.Button(frame, text=text, command=command)
        self.actions.append(b)
        return b

    # ── 파일 선택 ──────────────────────────────────────────
    def pick_dir(self, var):
        from tkinter import filedialog
        p = filedialog.askdirectory(initialdir=var.get() or None, parent=self.root)
        if p:
            var.set(str(Path(p)))

    def pick_txt_dir(self):
        from tkinter import filedialog
        p = filedialog.askdirectory(title="txt 13개가 들어 있는 폴더", parent=self.root)
        if not p:
            return
        folder = Path(p)
        self.a_txt.set(str(folder))
        # 결과는 기본으로 txt 폴더의 상위(접수번호 폴더)에 저장
        self.a_out.set(str(folder.parent if folder.name.lower() == "files" else folder))
        try:
            files = engine.find_txt_files(folder)
            self.write(f"txt {len(files)}개 확인: {folder}")
        except BotError as e:
            self.write(f"! {e}", "warn")

    def pick_03(self):
        from tkinter import filedialog, messagebox
        p = filedialog.askopenfilename(title="PT/SOC 입력을 마친 _03 파일",
                                       filetypes=[("엑셀 파일", "*.xlsx")], parent=self.root)
        if not p:
            return
        path = Path(p)
        self.b_xlsx.set(str(path))
        self.b_out.set(str(path.parent))
        info = engine.parse_out_name(path)
        if info:
            self.b_product.set(info["product"])
            if info["no"] != "03":
                self.write(f"! 선택한 파일 번호가 _{info['no']} 입니다. _03 파일이 맞는지 확인하세요.", "warn")
        try:
            names = engine.drug_names_in(path)
        except Exception as e:  # noqa: BLE001
            logging.exception("_03 파일 읽기 실패")
            messagebox.showerror("_03 파일 확인", explain(e) + f"\n\n로그: {self.log_path}", parent=self.root)
            return
        self.b_drug_box["values"] = names
        self.b_drug.set(names[0] if names else "")
        if not self.b_prefix.get() and names:
            self.b_prefix.set(re.split(r"\d", names[0], maxsplit=1)[0] or names[0])
        self.write(f"_03 파일 선택: {path.name}  /  의약품명: {', '.join(names) or '(없음)'}")

    # ── 실행 ───────────────────────────────────────────────
    def check_date(self, value):
        if not re.fullmatch(r"\d{8}", value):
            raise BotError("파일 날짜는 YYYYMMDD 8자리 숫자로 입력하세요.")

    def run_a(self):
        txt, product = self.a_txt.get().strip(), self.a_product.get().strip()
        drug = self.a_drug.get().strip() or product
        out, date = self.a_out.get().strip(), self.a_date.get().strip()
        try:
            if not txt: raise BotError("txt 폴더를 선택하세요.")
            if not product: raise BotError("제품명을 입력하세요.")
            if not out: raise BotError("결과 저장 폴더를 선택하세요.")
            self.check_date(date)
        except BotError as e:
            return self.show_error(e)
        self.start("A단계 실행 중…", "A", lambda: engine.run_step_a(Path(txt), Path(out), product, drug, date))

    def run_b(self):
        xlsx, product = self.b_xlsx.get().strip(), self.b_product.get().strip()
        drug, prefix = self.b_drug.get().strip(), self.b_prefix.get().strip()
        out, date = self.b_out.get().strip(), self.b_date.get().strip()
        try:
            if not xlsx: raise BotError("_03 파일을 선택하세요.")
            if not product: raise BotError("제품명을 입력하세요.")
            if not drug: raise BotError("의심/병용으로 나눌 의약품명이 없습니다. _03 파일 DRUG 시트 [의약품명]을 확인하세요.")
            if not prefix: raise BotError("시트 이름 앞부분을 입력하세요.")
            if not out: raise BotError("결과 저장 폴더를 선택하세요.")
            self.check_date(date)
        except BotError as e:
            return self.show_error(e)
        self.start("B단계 실행 중…", "B",
                   lambda: engine.run_step_b(Path(xlsx), Path(out), product, drug, prefix, date))

    def start(self, status, kind, fn):
        self.busy = True
        for w in self.actions:
            w.configure(state="disabled")
        self.status.set(status); self.progress.start(12)
        self.write("─" * 60)

        def work():
            try:
                self.events.put((kind, fn()))
            except Exception as exc:  # noqa: BLE001
                logging.exception("작업 실패")
                self.events.put(("error", exc))
        threading.Thread(target=work, daemon=True).start()

    def finish(self):
        self.busy = False
        self.progress.stop()
        for w in self.actions:
            w.configure(state="normal")
        self.b_drug_box.configure(state="readonly")

    def poll(self):
        from tkinter import messagebox
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    level, msg = value
                    self.write(msg, "warn" if level >= logging.WARNING else None)
                    continue
                self.finish()
                if kind == "error":
                    self.status.set("작업 실패 — 진행 기록과 로그를 확인하세요.")
                    self.show_error(value)
                    continue
                self.last_out = value.parent
                self.open_out_button.configure(state="normal")
                if kind == "A":
                    self.b_xlsx.set(str(value)); self.b_out.set(str(value.parent))
                    self.b_product.set(self.a_product.get().strip())
                    drug = self.a_drug.get().strip() or self.a_product.get().strip()
                    self.b_drug_box["values"] = [drug]; self.b_drug.set(drug)
                    if not self.b_prefix.get():
                        self.b_prefix.set(re.split(r"\d", drug, maxsplit=1)[0] or drug)
                    self.write(f"A단계 완료: {value}", "done")
                    self.status.set("A단계 완료 — _03 파일에 입력 후 B 탭에서 실행하세요.")
                    if messagebox.askyesno("A단계 완료",
                                           f"{value.name}\n\n_03 파일 EVENT 시트에 PT_한국어, SOC_한국어를 입력하고\n저장한 뒤 B 탭에서 실행하세요.\n\n"
                                           "지금 _03 파일을 열까요?", parent=self.root):
                        open_path(value)
                else:
                    self.write(f"B단계 완료: {value.parent}", "done")
                    self.status.set(f"B단계 완료 — {value.parent}")
                    if messagebox.askyesno("B단계 완료", f"_10, _20, _30 파일을 만들었습니다.\n{value.parent}\n\n_30 파일을 확인하며 LLT_한국어, 반영여부를 입력하세요.\n결과 폴더를 열까요?",
                                           parent=self.root):
                        open_path(value.parent)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    # ── 표시 ───────────────────────────────────────────────
    def write(self, msg, tag=None):
        if tag is None and msg.lstrip().startswith(("!", "※")):
            tag = "warn"
        self.text.configure(state="normal")
        self.text.insert("end", msg + "\n", tag or ())
        self.text.see("end")
        self.text.configure(state="disabled")

    def show_error(self, exc):
        from tkinter import messagebox
        msg = explain(exc)
        logging.error(msg.replace("\n", " / "))
        self.write(f"[오류] {msg}", "error")
        messagebox.showerror("오류", f"{msg}\n\n로그: {self.log_path}", parent=self.root)

    def close(self):
        from tkinter import messagebox
        if self.busy and not messagebox.askyesno("작업 중", "작업이 진행 중입니다. 종료할까요?", parent=self.root):
            return
        self.root.destroy()


# ════════════════════════════════════════════════════════════════════
# 실행
# ════════════════════════════════════════════════════════════════════
def run_cli(args) -> int:
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))
    date = args.date or dt.date.today().strftime("%Y%m%d")
    try:
        if args.step == "A":
            if not (args.txt and args.product):
                raise BotError("--step A 에는 --txt, --product 가 필요합니다.")
            out = args.out or (args.txt.parent if args.txt.name.lower() == "files" else args.txt)
            engine.run_step_a(args.txt, out, args.product, args.drug_name or args.product, date)
        else:
            if not (args.xlsx and args.sheet_prefix):
                raise BotError("--step B 에는 --xlsx, --sheet-prefix 가 필요합니다.")
            info = engine.parse_out_name(args.xlsx) or {}
            product = args.product or info.get("product")
            if not product:
                raise BotError("--product 를 입력하세요.")
            names = engine.drug_names_in(args.xlsx)
            drug = args.drug_name or (names[0] if len(names) == 1 else None)
            if not drug:
                raise BotError(f"--drug-name 을 입력하세요. DRUG 시트 [의약품명]: {names}")
            engine.run_step_b(args.xlsx, args.out or args.xlsx.parent, product, drug, args.sheet_prefix, date)
    except Exception as exc:  # noqa: BLE001
        logging.exception("작업 실패")
        print(f"[오류] {explain(exc)}")
        return 1
    print("완료")
    return 0


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(description=VERSION)
    parser.add_argument("--step", choices=["A", "B"])
    parser.add_argument("--txt", type=Path); parser.add_argument("--xlsx", type=Path)
    parser.add_argument("--product"); parser.add_argument("--drug-name"); parser.add_argument("--sheet-prefix")
    parser.add_argument("--out", type=Path); parser.add_argument("--date")
    args = parser.parse_args()
    log_path = setup_logging()
    if args.step:
        return run_cli(args)
    import tkinter as tk
    root = tk.Tk()
    App(root, log_path)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
