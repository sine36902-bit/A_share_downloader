"""Python 3.11+；安装：python -m pip install -U akshare pandas openpyxl requests"""
import math
import os
import queue
import re
import tempfile
import threading
import time
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import akshare as ak
import pandas as pd
import requests

ADJUST = {"不复权": "", "前复权": "qfq", "后复权": "hfq"}
EMPTY = "未获得交易记录：可能代码不存在、尚未上市、区间无交易或接口返回空数据。无法仅凭空结果区分原因；未生成文件。"


def parse_date(value):
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
        raise ValueError("日期格式必须为 YYYY-MM-DD。")
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError("日期不存在，请检查年月日。") from exc


def download(code, start, end, adjust, report):
    for attempt in range(3):
        report(f"正在下载（第 {attempt + 1}/3 次尝试）……")
        try:
            raw = ak.stock_zh_a_hist(
                symbol=code, period="daily", start_date=start.strftime("%Y%m%d"),
                end_date=end.strftime("%Y%m%d"), adjust=adjust,
            )
            break
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError) as exc:
            if attempt == 2:
                raise RuntimeError(f"网络请求失败，已尝试 3 次：{exc}") from exc
            report("网络异常，稍后重试……")
            time.sleep(2 * (attempt + 1))
    if raw is None or raw.empty:
        raise ValueError(EMPTY)
    if not {"日期", "收盘"}.issubset(raw.columns):
        raise ValueError("接口缺少“日期”或“收盘”字段，可能发生变更；停止导出。")
    data = raw[["日期", "收盘"]].rename(columns={"收盘": "收盘价"}).copy()
    data["日期"] = pd.to_datetime(data["日期"], errors="raise").dt.date
    data["收盘价"] = pd.to_numeric(data["收盘价"], errors="raise")
    if data.isna().any().any() or not data["收盘价"].map(math.isfinite).all():
        raise ValueError("接口返回缺失日期或无效价格，停止导出；不填补数据。")
    data = data[data["日期"].between(start, end)]
    if data.empty:
        raise ValueError(EMPTY)
    # 重复日期价格冲突时，不自行选择或估算价格。
    if (data.groupby("日期")["收盘价"].nunique() > 1).any():
        raise ValueError("同一日期出现不同收盘价，数据存在冲突；停止导出。")
    return data.drop_duplicates("日期").sort_values("日期").reset_index(drop=True)


def save_excel(data, path):
    # 先完整写入临时文件，避免导出失败破坏已有文件。
    fd, name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    try:
        with pd.ExcelWriter(name, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
            data.to_excel(writer, index=False, sheet_name="历史收盘价")
            sheet = writer.sheets["历史收盘价"]
            sheet.column_dimensions["A"].width = 14
            sheet.column_dimensions["B"].width = 16
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


class App:
    def __init__(self, root):
        self.root = root
        root.title("A 股历史收盘价下载器")
        self.events = queue.Queue()
        self.code = tk.StringVar(value="600519")
        self.start = tk.StringVar(value="2014-01-01")
        self.end = tk.StringVar(value="2023-12-31")
        self.adjust = tk.StringVar(value="不复权")
        self.folder = tk.StringVar()
        self.status = tk.StringVar(value="请选择保存目录，然后下载。")
        box = ttk.Frame(root, padding=16)
        box.grid(sticky="nsew")
        for row, (label, variable) in enumerate([
            ("股票代码", self.code), ("开始日期 YYYY-MM-DD", self.start),
            ("结束日期 YYYY-MM-DD", self.end),
        ]):
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Entry(box, textvariable=variable, width=28).grid(row=row, column=1, sticky="ew")
        ttk.Label(box, text="复权方式").grid(row=3, column=0, sticky="w")
        ttk.Combobox(box, textvariable=self.adjust, values=list(ADJUST),
                     state="readonly").grid(row=3, column=1, sticky="ew", pady=5)
        ttk.Label(box, text="保存目录").grid(row=4, column=0, sticky="w")
        ttk.Entry(box, textvariable=self.folder, state="readonly", width=36).grid(row=4, column=1)
        ttk.Button(box, text="选择…", command=self.choose_folder).grid(row=4, column=2, padx=5)
        self.button = ttk.Button(box, text="下载并导出", command=self.start_download)
        self.button.grid(row=5, column=0, columnspan=3, pady=12)
        ttk.Label(box, textvariable=self.status, wraplength=490).grid(row=6, column=0, columnspan=3, sticky="w")
        root.after(100, self.poll)

    def choose_folder(self):
        folder = filedialog.askdirectory(parent=self.root, title="选择保存目录")
        if folder:
            self.folder.set(folder)

    def start_download(self):
        try:
            code = self.code.get().strip()  # 始终使用字符串，保留前导零。
            if not re.fullmatch(r"[0-9]{6}", code):
                raise ValueError("股票代码必须为六位数字，例如 000001；不含交易所前缀。")
            start, end = parse_date(self.start.get().strip()), parse_date(self.end.get().strip())
            if start > end:
                raise ValueError("开始日期不能晚于结束日期。")
            folder = Path(self.folder.get())
            if not self.folder.get() or not folder.is_dir():
                raise ValueError("请先选择有效的保存目录。")
            label = self.adjust.get()
            path = folder / f"{code}_{start:%Y%m%d}_{end:%Y%m%d}_{label}.xlsx"
        except ValueError as exc:
            messagebox.showerror("输入错误", str(exc), parent=self.root)
            return
        self.button.state(["disabled"])
        self.status.set("正在连接东方财富……")
        # 仅一个后台线程负责下载，让窗口保持响应；所有界面操作在主线程执行。
        def worker():
            try:
                data = download(code, start, end, ADJUST[label],
                                lambda msg: self.events.put(("status", msg)))
                self.events.put(("data", (data, path)))
            except Exception as exc:
                self.events.put(("error", f"{type(exc).__name__}: {exc}"))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "status":
                    self.status.set(value)
                elif kind == "error":
                    self.button.state(["!disabled"])
                    self.status.set("下载失败，未导出文件。")
                    messagebox.showerror("下载失败", value, parent=self.root)
                else:
                    data, path = value
                    try:
                        if path.exists() and not messagebox.askyesno(
                            "确认覆盖", f"文件已存在，是否覆盖？\n{path}", parent=self.root
                        ):
                            self.status.set("已取消导出。")
                            continue
                        self.status.set("正在导出 Excel……")
                        self.root.update_idletasks()
                        save_excel(data, path)
                        self.status.set(f"成功：{len(data)} 行；{path}")
                        messagebox.showinfo("导出成功", f"实际获得 {len(data)} 行数据。\n保存路径：{path}", parent=self.root)
                    except Exception as exc:
                        self.status.set("导出失败。")
                        messagebox.showerror("导出失败", f"请检查目录权限、磁盘空间，以及文件是否被 Excel 占用。\n{type(exc).__name__}: {exc}", parent=self.root)
                    finally:
                        self.button.state(["!disabled"])
        except queue.Empty:
            pass
        self.root.after(100, self.poll)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
