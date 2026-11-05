"""Python 3.11+：pip install -U akshare pandas openpyxl requests
行情仅来自东方财富；名称查询失败不阻止行情。后台串行请求，无自行聚合或补值。
"""
from __future__ import annotations
import math
import os
import queue
import re
import tempfile
import threading
import time
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import akshare as ak
import pandas as pd
import requests

PERIODS = {"日线": "daily", "周线": "weekly", "月线": "monthly"}
ADJUSTS = {"不复权": "", "前复权": "qfq", "后复权": "hfq"}
FIELDS = ["日期", "开盘价", "最高价", "最低价", "收盘价", "成交量", "成交额", "涨跌幅", "换手率"]
SOURCE_FIELDS = dict(zip(FIELDS, ["日期", "开盘", "最高", "最低", "收盘", "成交量", "成交额", "涨跌幅", "换手率"]))
NETWORK_ERRORS = (requests.exceptions.RequestException, TimeoutError, ConnectionError)
try:  # 新版 AKShare 的 HTTP 实现可能使用 curl_cffi。
    from curl_cffi.requests.exceptions import RequestException as CurlError
    NETWORK_ERRORS += (CurlError,)
except ImportError:
    pass


class NoData(ValueError):
    pass


def parse_codes(text):
    result = []
    for token in re.split(r"[,，;；\s]+", text.strip()):
        if not token:
            continue
        match = re.fullmatch(r"(?:(股票|指数)[:：])?([0-9]{6})", token)
        if not match:
            raise ValueError(f"代码格式错误：{token}；请使用 600519 或 指数:000300。")
        key = (match[1] or "股票", match[2])
        if key not in result:
            result.append(key)
    if not result:
        raise ValueError("请至少输入一只股票或指数。")
    return result


def parse_date(value):
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
        raise ValueError("日期必须为 YYYY-MM-DD。")
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"无效日期：{value}") from exc


def retry_call(func, report, **kwargs):
    for attempt in range(3):
        try:
            # 不传 timeout：stock_zh_a_hist 的公开签名不支持此参数。
            return func(**kwargs)
        except NETWORK_ERRORS:
            if attempt == 2:
                raise
            report(f"网络异常，{2 * (attempt + 1)} 秒后重试（最多 3 次）")
            time.sleep(2 * (attempt + 1))


def load_names(kind, report):
    if kind == "股票":
        frame = retry_call(ak.stock_info_a_code_name, report)
        a, b = "code", "name"
    else:
        frame = retry_call(ak.index_stock_info, report)
        a, b = "index_code", "display_name"
    return {str(code).zfill(6): str(name) for code, name in zip(frame[a], frame[b])
            if pd.notna(name)}


def column_name(field, kind):
    if field in ("开盘价", "最高价", "最低价", "收盘价"):
        return field + ("（元）" if kind == "股票" else "（点）")
    return field + {"成交量": "（手）", "成交额": "（元）", "涨跌幅": "（%）", "换手率": "（%）"}.get(field, "")


def clean_data(raw, kind, config):
    if raw is None or raw.empty:
        raise NoData("接口返回空数据：可能代码不存在、区间无交易、尚未上市或接口限制；不能仅凭空结果确定原因。")
    if not {"日期", "收盘"}.issubset(raw.columns):
        raise ValueError("接口缺少必需字段“日期”或“收盘”，停止使用该结果。")
    # 对所有已知字段检查重复冲突，而非只检查选中字段。
    columns = [v for v in SOURCE_FIELDS.values() if v in raw.columns]
    data = raw[columns].copy()
    data["日期"] = pd.to_datetime(data["日期"], errors="raise").dt.date
    if data["日期"].isna().any():
        raise ValueError("接口存在缺失日期。")
    data = data[data["日期"].between(config["start"], config["end"])].copy()
    if data.empty:
        raise NoData("接口没有返回请求区间内的记录；不生成行情表。")
    for col in columns[1:]:
        data[col] = pd.to_numeric(data[col], errors="raise")
        if data[col].dropna().map(lambda x: not math.isfinite(x)).any():
            raise ValueError(f"字段 {col} 含无穷值。")
    if data["收盘"].isna().any():
        raise ValueError("收盘价存在缺失，停止导出，不补值。")
    if (data.groupby("日期").nunique(dropna=False) > 1).any().any():
        raise ValueError("同一日期存在冲突记录；未擅自删除，请核对数据源。")
    before = len(data)
    data = data.drop_duplicates("日期").sort_values("日期").reset_index(drop=True)
    notes = []
    if before > len(data):
        notes.append(f"去除 {before - len(data)} 条完全相同的重复日期记录")
    selected = []
    for field in config["fields"]:
        source = SOURCE_FIELDS[field]
        # 指数接口即使出现占位换手率，也不作为有效换手率导出。
        if source not in data or (kind == "指数" and field == "换手率"):
            notes.append(f"跳过不支持字段：{field}")
        elif field not in ("日期", "收盘价") and data[source].isna().all():
            notes.append(f"跳过全空字段：{field}")
        else:
            selected.append(field)
            if data[source].isna().any():
                notes.append(f"{field} 部分缺失，原样留空")
    result = data[[SOURCE_FIELDS[f] for f in selected]].copy()
    result.columns = [column_name(f, kind) for f in selected]
    return result, notes


def fetch_one(key, config, name, name_note, report):
    kind, code = key
    interface = "stock_zh_a_hist" if kind == "股票" else "index_zh_a_hist"
    notes = [name_note] if name_note else []
    notes.append("仅代表接口实际返回记录，未核验完整交易日覆盖；不补齐缺失数据")
    if kind == "指数":
        notes.append("指数按点位、不复权口径；忽略个股复权选择；接口仅承诺近期数据，不能保证完整历史")
    if config["period"] != "日线":
        notes.append("源站原生周/月线，无本地聚合；按源站周期日期筛选，边界周期可能跨请求起点，当前周期可能尚未结束")
    record = {"证券类型": kind, "代码": code, "证券名称": name or code,
              "数据来源": "东方财富（AKShare）", "接口名称": interface,
              "请求开始日期": str(config["start"]), "请求结束日期": str(config["end"]),
              "实际开始日期": "", "实际结束日期": "", "行情周期": config["period"],
              "复权或指数口径": config["adjust"] if kind == "股票" else "指数点位（不适用个股复权）",
              "下载时间": "", "状态": "失败", "记录数": 0, "说明": "",
              "AKShare版本": getattr(ak, "__version__", "未知"),
              "名称来源": "stock_info_a_code_name" if kind == "股票" else "index_stock_info（聚宽）"}
    data = None
    try:
        kwargs = dict(symbol=code, period=PERIODS[config["period"]],
                      start_date=config["start"].strftime("%Y%m%d"), end_date=config["end"].strftime("%Y%m%d"))
        if kind == "股票":
            kwargs["adjust"] = ADJUSTS[config["adjust"]]
        raw = retry_call(getattr(ak, interface), report, **kwargs)
        data, clean_notes = clean_data(raw, kind, config)
        notes.extend(clean_notes)
        record.update({"状态": "成功", "记录数": len(data),
                       "实际开始日期": str(data["日期"].iloc[0]), "实际结束日期": str(data["日期"].iloc[-1])})
    except NoData as exc:
        record["状态"] = "无数据"
        notes.append(str(exc))
    except Exception as exc:
        notes.append(f"{type(exc).__name__}: {exc}")
    record["下载时间"] = datetime.now().astimezone().isoformat(timespec="seconds")
    record["说明"] = "；".join(notes)
    return {"key": key, "meta": record, "data": data}


def safe_name(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f\[\]]', "_", str(value)).strip(" .'") or "未命名"


def sheet_names(results):
    used = {"下载说明"}
    names = {}
    for item in results:
        if item["data"] is None:
            continue
        m = item["meta"]
        base = safe_name(f'{m["证券类型"]}_{m["代码"]}_{m["证券名称"]}')[:31]
        name, i = base, 1
        while name.casefold() in used:
            suffix = f"_{i}"
            name = base[:31-len(suffix)] + suffix
            i += 1
        used.add(name.casefold())
        names[item["key"]] = name
    return names


def export_plan(results, config, folder, fmt):
    start, end = config["start"].strftime("%Y%m%d"), config["end"].strftime("%Y%m%d")
    good = [x for x in results if x["data"] is not None]
    def stem(item):
        m = item["meta"]
        adjust = "_" + config["adjust"] if m["证券类型"] == "股票" else ""
        return safe_name(f'{m["证券类型"]}_{m["代码"]}_{m["证券名称"]}_{config["period"]}{adjust}_{start}_{end}')[:180]
    if fmt == "Excel":
        name = stem(good[0]) if len(results) == 1 else f"A股历史行情_{start}_{end}_{config['period']}"
        return [(folder / (name + ".xlsx"), None)]
    return [(folder / (stem(x) + ".csv"), x) for x in good] + [(folder / "下载说明.csv", None)]


def atomic_write(path, writer):
    fd, temporary = tempfile.mkstemp(suffix=path.suffix, dir=path.parent)
    os.close(fd)
    try:
        writer(Path(temporary))
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def write_exports(results, plan, fmt):
    """返回每个文件的写入结果；单个 CSV 写入失败不阻止其他文件。"""
    outcomes = []
    meta = [dict(x["meta"], 导出文件="", 导出状态="未导出") for x in results]
    by_key = {x["key"]: m for x, m in zip(results, meta)}
    if fmt == "Excel":
        path = plan[0][0]
        names = sheet_names(results)
        for item in results:
            if item["data"] is not None:
                by_key[item["key"]].update(导出文件=f"{path.name} / {names[item['key']]}", 导出状态="成功")
        def write_book(temp):
            with pd.ExcelWriter(temp, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
                for item in results:
                    if item["data"] is not None:
                        item["data"].to_excel(writer, sheet_name=names[item["key"]], index=False)
                pd.DataFrame(meta).to_excel(writer, sheet_name="下载说明", index=False)
                for sheet in writer.book:
                    sheet.freeze_panes = "A2"
                    for cells in sheet.columns:
                        sheet.column_dimensions[cells[0].column_letter].width = 22
                    # 元数据包含外部名称/错误信息：强制文本，避免被当成公式。
                    for row in sheet:
                        for cell in row:
                            if isinstance(cell.value, str):
                                cell.data_type = "s"
        try:
            atomic_write(path, write_book)
            outcomes.append(f"成功：{path}")
        except Exception as exc:
            outcomes.append(f"失败：{path}：{type(exc).__name__}: {exc}")
    else:
        for path, item in plan:
            try:
                frame = item["data"] if item is not None else pd.DataFrame(meta)
                atomic_write(path, lambda temp: frame.to_csv(temp, index=False, encoding="utf-8-sig"))
                outcomes.append(f"成功：{path}")
                if item is not None:
                    by_key[item["key"]].update(导出文件=path.name, 导出状态="成功")
            except Exception as exc:
                error = f"{type(exc).__name__}: {exc}"
                outcomes.append(f"失败：{path}：{error}")
                if item is not None:
                    by_key[item["key"]].update(导出文件=path.name, 导出状态="失败：" + error)
    return outcomes


class App:
    def __init__(self, root):
        self.root, self.events = root, queue.Queue()
        self.results, self.config, self.busy = {}, None, False
        root.title("极简 A 股历史行情下载器 V2")
        root.geometry("1050x760")
        main = ttk.Frame(root, padding=12)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text="证券代码（逗号、空格、分号、换行分隔；指数须加“指数:”）").pack(anchor="w")
        self.codes = tk.Text(main, height=3)
        self.codes.insert("1.0", "600519, 股票:000001, 指数:000300, 指数:000001")
        self.codes.pack(fill="x")
        conditions = ttk.Frame(main)
        conditions.pack(fill="x", pady=8)
        self.start, self.end = tk.StringVar(value="2014-01-01"), tk.StringVar(value=str(date.today()))
        self.period, self.adjust = tk.StringVar(value="日线"), tk.StringVar(value="不复权")
        for label, var in [("开始日期", self.start), ("结束日期", self.end)]:
            ttk.Label(conditions, text=label).pack(side="left", padx=4)
            ttk.Entry(conditions, textvariable=var, width=12).pack(side="left")
        for label, var, choices in [("周期", self.period, PERIODS), ("复权", self.adjust, ADJUSTS)]:
            ttk.Label(conditions, text=label).pack(side="left", padx=8)
            ttk.Combobox(conditions, textvariable=var, values=list(choices), state="readonly", width=9).pack(side="left")
        fields = ttk.Frame(main)
        fields.pack(fill="x")
        self.fields = {}
        for field in FIELDS:
            var = tk.BooleanVar(value=field in ("日期", "收盘价"))
            self.fields[field] = var
            check = ttk.Checkbutton(fields, text=field, variable=var)
            check.pack(side="left")
            if field in ("日期", "收盘价"):
                check.state(["disabled"])
        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=10)
        self.fetch_btn = ttk.Button(actions, text="获取并预览", command=self.fetch)
        self.fetch_btn.pack(side="left")
        self.retry_btn = ttk.Button(actions, text="重新获取失败/无数据项目", command=lambda: self.fetch(True))
        self.retry_btn.pack(side="left", padx=6)
        self.format = tk.StringVar(value="Excel")
        ttk.Combobox(actions, textvariable=self.format, values=["Excel", "CSV"], width=7, state="readonly").pack(side="left")
        self.export_btn = ttk.Button(actions, text="导出数据", command=self.export)
        self.export_btn.pack(side="right")
        paths = ttk.Frame(main)
        paths.pack(fill="x")
        self.folder = tk.StringVar()
        ttk.Button(paths, text="选择导出目录", command=self.choose_folder).pack(side="left")
        ttk.Entry(paths, textvariable=self.folder, state="readonly").pack(side="left", fill="x", expand=True)
        self.status = tk.StringVar(value="等待获取。字段、日期等条件变更后，请重新获取；导出使用预览时的快照。")
        ttk.Label(main, textvariable=self.status, wraplength=990).pack(anchor="w", pady=6)
        self.listbox = tk.Listbox(main, height=5, exportselection=False)
        self.listbox.pack(fill="x")
        self.listbox.bind("<<ListboxSelect>>", self.preview)
        self.detail = tk.StringVar()
        ttk.Label(main, textvariable=self.detail, wraplength=990).pack(anchor="w", pady=5)
        table = ttk.Frame(main)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, show="headings", height=10)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=scroll.set)
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        root.after(100, self.poll)

    def choose_folder(self):
        folder = filedialog.askdirectory(parent=self.root)
        if folder:
            self.folder.set(folder)

    def set_busy(self, value):
        self.busy = value
        for button in (self.fetch_btn, self.retry_btn, self.export_btn):
            button.state(["disabled" if value else "!disabled"])

    def fetch(self, retry=False):
        if self.busy:
            return
        try:
            if retry:
                keys = [key for key, item in self.results.items() if item["data"] is None]
                if not keys or self.config is None:
                    raise ValueError("没有需要重试的项目。")
                config = self.config
            else:
                keys = parse_codes(self.codes.get("1.0", "end"))
                start, end = parse_date(self.start.get().strip()), parse_date(self.end.get().strip())
                if start > end:
                    raise ValueError("开始日期不能晚于结束日期。")
                config = dict(start=start, end=end, period=self.period.get(), adjust=self.adjust.get(),
                              fields=[f for f in FIELDS if self.fields[f].get()])
                self.results = {}
                self.config = config
                self.refresh()
        except ValueError as exc:
            messagebox.showerror("输入提示", str(exc), parent=self.root)
            return
        self.set_busy(True)
        self.status.set("开始获取；股票和指数使用独立接口。")
        def worker():
            names = {}
            try:
                for i, key in enumerate(keys, 1):
                    kind, code = key
                    def report(msg):
                        self.events.put(("status", f"正在获取第 {i}/{len(keys)} 只：{kind}:{code}；{msg}"))
                    report("获取行情与名称……")
                    if kind not in names:
                        try:
                            names[kind] = (load_names(kind, report), "")
                        except Exception as exc:
                            names[kind] = ({}, f"名称获取失败：{type(exc).__name__}: {exc}")
                    mapping, error = names[kind]
                    name = mapping.get(code, "")
                    note = "" if name else error or "名称表中未找到代码；名称获取失败，保留代码"
                    item = fetch_one(key, config, name, note, report)
                    self.events.put(("result", item))
            except Exception as exc:
                self.events.put(("error", str(exc)))
            finally:
                self.events.put(("done", None))
        threading.Thread(target=worker, daemon=True).start()

    def refresh(self):
        self.listbox.delete(0, "end")
        for item in self.results.values():
            m = item["meta"]
            self.listbox.insert("end", f'{m["证券类型"]}:{m["代码"]}  {m["证券名称"]}  {m["状态"]}  {m["记录数"]} 条')
        if self.results:
            self.listbox.selection_set(0)
        self.preview()

    def preview(self, event=None):
        self.tree.delete(*self.tree.get_children())
        selected = self.listbox.curselection()
        if not selected:
            self.tree["columns"] = ()
            self.detail.set("")
            return
        item = list(self.results.values())[selected[0]]
        m = item["meta"]
        self.detail.set(f'{m["证券类型"]} {m["证券名称"]} ({m["代码"]}) | {m["记录数"]} 条 | {m["行情周期"]} | {m["复权或指数口径"]}\n{m["说明"]}')
        data = item["data"]
        if data is None:
            self.tree["columns"] = ()
            return
        self.tree["columns"] = list(data.columns)
        for col in data:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=125, stretch=False)
        for row in data.head(10).itertuples(index=False, name=None):
            self.tree.insert("", "end", values=["" if pd.isna(v) else str(v) for v in row])

    def export(self):
        if self.busy:
            return
        results = list(self.results.values())
        if not any(x["data"] is not None for x in results):
            messagebox.showwarning("无法导出", "没有成功获取的行情，请先获取或重试。", parent=self.root)
            return
        if not self.folder.get():
            self.choose_folder()
        folder = Path(self.folder.get())
        if not self.folder.get() or not folder.is_dir():
            return
        fmt = self.format.get()
        plan = export_plan(results, self.config, folder, fmt)
        existing = [str(path) for path, _ in plan if path.exists()]
        if existing and not messagebox.askyesno("确认覆盖", "以下文件已存在，是否覆盖？\n" + "\n".join(existing), parent=self.root):
            return
        self.set_busy(True)
        self.status.set("正在导出预览数据，不重新请求行情……")
        def worker():
            try:
                self.events.put(("exported", write_exports(results, plan, fmt)))
            except Exception as exc:
                self.events.put(("error", f"导出失败：{exc}"))
                self.events.put(("done", None))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "status":
                    self.status.set(value)
                elif kind == "result":
                    self.results[value["key"]] = value
                    self.refresh()
                elif kind == "error":
                    messagebox.showerror("错误", value, parent=self.root)
                elif kind == "done":
                    self.set_busy(False)
                    good = sum(x["data"] is not None for x in self.results.values())
                    self.status.set(f"获取结束：成功 {good}，失败/无数据 {len(self.results)-good}；选择列表查看详情，可重试失败项目。")
                elif kind == "exported":
                    self.set_busy(False)
                    failed = sum(x.startswith("失败") for x in value)
                    self.status.set(f"导出结束：成功 {len(value)-failed} 个文件，失败 {failed} 个文件。")
                    show = messagebox.showerror if failed else messagebox.showinfo
                    show("导出结果", "\n".join(value) + ("\n请检查目录权限、磁盘空间及文件是否被 Excel 占用。" if failed else ""), parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
