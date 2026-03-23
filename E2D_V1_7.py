import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import os

try:
    import ttkbootstrap as tb
except Exception:
    tb = None

from ui_config import (
    WINDOW_SIZE,
    BG_COLOR,
    PRIMARY_COLOR,
    SECONDARY_TEXT,
    PLACEHOLDER_COLOR,
    LABEL_FONT,
    BUTTON_FONT,
    ENTRY_FONT,
    apply_design_system,
)
from ui.widgets import make_button, make_date_entry, set_date_value
from core.state import AppState
from ui.components import ProductLineManager
from core.mapping import ColumnMapper
from core.processing import ExcelProcessor
from core.process_impl import process_raw_excel, get_sheets_with_data


def _create_toast(root, message, duration=1800):
    try:
        toast = tk.Toplevel(root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(bg="#9CA3AF")
        label = tk.Label(toast, text=message, fg="white", bg="#9CA3AF", font=("Microsoft YaHei UI", 10))
        label.pack(padx=12, pady=8)
        tw = label.winfo_reqwidth() + 24
        th = label.winfo_reqheight() + 16
        root.update_idletasks()
        rx = root.winfo_rootx()
        ry = root.winfo_rooty()
        rw = root.winfo_width()
        rh = root.winfo_height()
        toast.geometry(f"{tw}x{th}+{rx + rw//2 - tw//2}+{ry + rh - th - 40}")
        toast.after(duration, toast.destroy)
    except Exception as e:
        logging.warning("Toast creation failed: %s", e)


def _build_column_mapping_tab(container, column_mapper):
    main_frame = ttk.Frame(container, padding=35)
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="列映射配置", font=("Microsoft YaHei UI", 18, "bold"),
              foreground=PRIMARY_COLOR).pack(pady=(0, 30))

    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

    mapping_frame = ttk.Frame(notebook)
    notebook.add(mapping_frame, text="列映射配置")

    output_frame = ttk.Frame(notebook)
    notebook.add(output_frame, text="输出列配置")

    mapping_canvas = tk.Canvas(mapping_frame, highlightthickness=0, bg=BG_COLOR)
    mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=mapping_canvas.yview)
    mapping_scrollable_frame = ttk.Frame(mapping_canvas, style='TFrame')
    mapping_scrollable_frame.bind("<Configure>",
                                   lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all")))
    mapping_window_id = mapping_canvas.create_window((0, 0), window=mapping_scrollable_frame, anchor="nw")
    mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
    mapping_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    mapping_scrollbar.pack(side="right", fill="y")
    mapping_frame.bind('<Configure>',
                       lambda e: mapping_canvas.itemconfigure(mapping_window_id, width=mapping_frame.winfo_width()))

    output_canvas = tk.Canvas(output_frame, highlightthickness=0, bg=BG_COLOR)
    output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_canvas.yview)
    output_scrollable_frame = ttk.Frame(output_canvas, style='TFrame')
    output_scrollable_frame.bind("<Configure>",
                                  lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all")))
    output_window_id = output_canvas.create_window((0, 0), window=output_scrollable_frame, anchor="nw")
    output_canvas.configure(yscrollcommand=output_scrollbar.set)
    output_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    output_scrollbar.pack(side="right", fill="y")
    output_frame.bind('<Configure>',
                      lambda e: output_canvas.itemconfigure(output_window_id, width=output_frame.winfo_width()))

    mapping_entries = {}
    output_entries = {}

    for i, (target, aliases) in enumerate(column_mapper.get_mapping().items()):
        ttk.Label(mapping_scrollable_frame, text=f"{target}:").grid(
            row=i, column=0, sticky="w", padx=(20, 20), pady=(15, 15))
        entry = ttk.Entry(mapping_scrollable_frame, width=40)
        entry.insert(0, ", ".join(aliases))
        entry.grid(row=i, column=1, sticky="ew", pady=(15, 15), padx=(0, 20))
        mapping_entries[target] = entry
    mapping_scrollable_frame.columnconfigure(1, weight=1)

    for i, (source, target) in enumerate(column_mapper.get_output_columns().items()):
        ttk.Label(output_scrollable_frame, text=f"{source}:").grid(
            row=i, column=0, sticky="w", padx=(20, 20), pady=(15, 15))
        entry = ttk.Entry(output_scrollable_frame, width=40)
        entry.insert(0, target)
        entry.grid(row=i, column=1, sticky="ew", pady=(15, 15), padx=(0, 20))
        output_entries[source] = entry
    output_scrollable_frame.columnconfigure(1, weight=1)

    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(30, 10))
    ttk.Frame(button_frame).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def save_config():
        new_mapping = {}
        for target, entry in mapping_entries.items():
            aliases = [alias.strip() for alias in entry.get().split(",") if alias.strip()]
            new_mapping[target] = aliases if aliases else [target]
        new_output_columns = {}
        for source, entry in output_entries.items():
            new_output_columns[source] = entry.get().strip() or source
        column_mapper.column_mapping = new_mapping
        column_mapper.output_columns = new_output_columns
        column_mapper.save_mapping()
        messagebox.showinfo("成功", "配置已保存！")

    def import_config():
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            column_mapper.load_from_path(path)
            for target, entry in mapping_entries.items():
                aliases = column_mapper.get_mapping().get(target, [target])
                entry.delete(0, tk.END)
                entry.insert(0, ", ".join(aliases))
            for source, entry in output_entries.items():
                target = column_mapper.get_output_columns().get(source, source)
                entry.delete(0, tk.END)
                entry.insert(0, target)
            messagebox.showinfo("成功", "配置导入成功")
        except Exception as e:
            messagebox.showerror("错误", f"配置导入失败: {e}")

    def export_config():
        path = filedialog.asksaveasfilename(defaultexension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            column_mapper.save_to_path(path)
            messagebox.showinfo("成功", "配置导出成功")
        except Exception as e:
            messagebox.showerror("错误", f"配置导出失败: {e}")

    def refresh_config():
        try:
            column_mapper.load_mapping()
            for target, entry in mapping_entries.items():
                aliases = column_mapper.get_mapping().get(target, [target])
                entry.delete(0, tk.END)
                entry.insert(0, ", ".join(aliases))
            for source, entry in output_entries.items():
                target = column_mapper.get_output_columns().get(source, source)
                entry.delete(0, tk.END)
                entry.insert(0, target)
            messagebox.showinfo("成功", "配置刷新成功")
        except Exception as e:
            messagebox.showerror("错误", f"配置刷新失败: {e}")

    ttk.Button(button_frame, text="导入配置", command=import_config, width=12).pack(side=tk.RIGHT, padx=(8, 0))
    ttk.Button(button_frame, text="导出配置", command=export_config, width=12).pack(side=tk.RIGHT, padx=(8, 0))
    ttk.Button(button_frame, text="刷新配置", command=refresh_config, width=12).pack(side=tk.RIGHT, padx=(8, 0))
    ttk.Button(button_frame, text="保存配置", command=save_config, width=15).pack(side=tk.RIGHT)


def _build_info_tab(info_tab):
    info_frame = ttk.Frame(info_tab, padding=30)
    info_frame.pack(fill=tk.BOTH, expand=True)
    ttk.Label(info_frame, text="工具作者：July-Chen-JIE", font=("Microsoft YaHei UI", 12)).pack(anchor=tk.W, pady=(0, 10))
    ttk.Label(info_frame, text="版本：V1.7", font=("Microsoft YaHei UI", 12)).pack(anchor=tk.W, pady=(0, 10))
    ttk.Label(info_frame, text="更新日期：20251128", font=("Microsoft YaHei UI", 12)).pack(anchor=tk.W)


def _show_processing_overlay(root, process_btn, exit_btn, progress_callback):
    overlay = tk.Toplevel(root)
    overlay.overrideredirect(True)
    overlay.attributes("-topmost", True)
    try:
        overlay.attributes("-alpha", 0.0)
    except Exception:
        pass
    try:
        overlay.grab_set()
    except Exception:
        pass
    x = root.winfo_rootx()
    y = root.winfo_rooty()
    w = root.winfo_width()
    h = root.winfo_height()
    overlay.geometry(f"{w}x{h}+{x}+{y}")
    mask = tk.Frame(overlay, bg="#000000")
    mask.pack(fill=tk.BOTH, expand=True)
    card = ttk.Frame(mask, padding=30)
    card.place(relx=0.5, rely=0.5, anchor="center")
    pvar = tk.DoubleVar(value=0)
    pstyle = ttk.Style()
    pstyle.configure("Custom.Horizontal.TProgressbar", troughcolor="#E5E7EB", background=PRIMARY_COLOR, borderwidth=0)
    pbar = ttk.Progressbar(card, variable=pvar, maximum=100, length=380,
                          style="Custom.Horizontal.TProgressbar")
    pbar.pack(pady=(0, 12))
    plabel = ttk.Label(card, text="准备处理...", font=LABEL_FONT)
    plabel.pack()

    import threading
    cancel_event = threading.Event()
    make_button(card, text="取消", command=lambda: cancel_event.set(), width=10, role="danger").pack(pady=(8, 0))

    def upd(p, m):
        pvar.set(p)
        plabel.config(text=m)
        overlay.update_idletasks()

    return overlay, upd, cancel_event


def create_app():
    root = tb.Window(themename='cosmo') if tb else tk.Tk()
    root.title("Excel数据处理工具 v1.7")
    root.geometry(WINDOW_SIZE)
    root.configure(bg=BG_COLOR)
    style = ttk.Style()
    apply_design_system(style)

    app_notebook = ttk.Notebook(root)
    app_notebook.pack(fill=tk.BOTH, expand=True)

    main_tab = ttk.Frame(app_notebook)
    app_notebook.add(main_tab, text="数据处理")

    mapping_tab = ttk.Frame(app_notebook)
    app_notebook.add(mapping_tab, text="产品线映射")

    column_tab = ttk.Frame(app_notebook)
    app_notebook.add(column_tab, text="列映射配置")

    info_tab = ttk.Frame(app_notebook)
    app_notebook.add(info_tab, text="软件信息")

    title_frame = tk.Frame(main_tab, bg=BG_COLOR)
    title_frame.pack(fill=tk.X, pady=(0, 16))
    title_container = tk.Frame(title_frame, bg=BG_COLOR)
    title_container.pack(side=tk.LEFT)
    ttk.Label(title_container, text="Excel数据处理工具", font=("Microsoft YaHei UI", 24, "bold"),
              foreground=PRIMARY_COLOR).pack(anchor=tk.W)
    ttk.Label(title_container, text="智能化数据处理和报表生成", font=("Microsoft YaHei UI", 12),
              foreground=SECONDARY_TEXT).pack(anchor=tk.W, pady=(8, 0))
    ttk.Label(title_frame, text="v1.7", font=("Microsoft YaHei UI", 12),
              foreground=PLACEHOLDER_COLOR).pack(side=tk.RIGHT, pady=(15, 0))

    input_entry = tk.StringVar()
    output_entry = tk.StringVar()
    start_date_var = tk.StringVar()
    end_date_var = tk.StringVar()
    replace_mode_var = tk.StringVar(value='overwrite')
    app_state = AppState()

    def load_app_state():
        try:
            with open('app_state.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
            iv = data.get('input_file', '')
            ov = data.get('output_dir', '')
            rm = data.get('replace_mode', 'overwrite')
            if iv:
                input_entry.set(iv)
                app_state.input_file = iv
            if ov:
                output_entry.set(ov)
                app_state.output_dir = ov
            replace_mode_var.set(rm)
        except Exception:
            pass

    def save_app_state():
        try:
            data = {
                'input_file': app_state.input_file or input_entry.get().strip(),
                'output_dir': app_state.output_dir or output_entry.get().strip(),
                'replace_mode': replace_mode_var.get(),
            }
            with open('app_state.json', 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    import json

    def select_input_file():
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            input_entry.set(path)
            output_entry.set(os.path.dirname(path))
            app_state.input_file = path
            app_state.output_dir = os.path.dirname(path)
            _create_toast(root, "已选择输入文件")
            save_app_state()

    def select_output_dir():
        path = filedialog.askdirectory()
        if path:
            output_entry.set(path)
            app_state.output_dir = path
            _create_toast(root, "已选择输出目录")
            save_app_state()

    file_frame = ttk.LabelFrame(main_tab, text="▌文件设置", padding=16)
    file_frame.pack(fill=tk.X, pady=(0, 25))
    file_frame.columnconfigure(1, weight=1)
    ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, pady=(8, 12))
    ttk.Entry(file_frame, textvariable=input_entry, width=42, font=ENTRY_FONT).grid(
        row=0, column=1, sticky=tk.EW, padx=(12, 12), pady=(8, 12))
    ttk.Button(file_frame, text="浏览", command=select_input_file).grid(row=0, column=2, pady=(8, 12))
    ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=(0, 12))
    ttk.Entry(file_frame, textvariable=output_entry, width=42, font=ENTRY_FONT).grid(
        row=1, column=1, sticky=tk.EW, padx=(12, 12), pady=(0, 12))
    ttk.Button(file_frame, text="浏览", command=select_output_dir).grid(row=1, column=2, pady=(0, 12))

    date_frame = ttk.LabelFrame(main_tab, text="▌日期筛选", padding=16)
    date_frame.pack(fill=tk.X, pady=(0, 25))
    ttk.Label(date_frame, text="起始日期:").grid(row=0, column=0, sticky=tk.W, pady=(8, 12))
    start_date_entry = make_date_entry(date_frame, width=14, dateformat="%Y/%m/%d",
                                      bootstyle="success", firstweekday=6)
    start_date_entry.grid(row=0, column=1, sticky=tk.W, padx=(12, 20), pady=(8, 12))
    ttk.Label(date_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, pady=(8, 12))
    end_date_entry = make_date_entry(date_frame, width=14, dateformat="%Y/%m/%d",
                                     bootstyle="success", firstweekday=6)
    end_date_entry.grid(row=0, column=3, sticky=tk.W, padx=(12, 0), pady=(8, 12))

    def set_start_date(val):
        start_date_var.set(val)
        app_state.start_date = val
        set_date_value(start_date_entry, val)

    def set_end_date(val):
        end_date_var.set(val)
        app_state.end_date = val
        set_date_value(end_date_entry, val)

    # 暂时移除事件绑定以避免崩溃
    # def on_start_date_change(event):
    #     try:
    #         widget = event.widget
    #         if hasattr(widget, 'get'):
    #             try:
    #                 val = widget.get()
    #                 set_start_date(val)
    #             except Exception:
    #                 pass
    #     except Exception as e:
    #         logging.warning("Start date change error: %s", e)

    # def on_end_date_change(event):
    #     try:
    #         widget = event.widget
    #         if hasattr(widget, 'get'):
    #             try:
    #                 val = widget.get()
    #                 set_end_date(val)
    #             except Exception:
    #                 pass
    #     except Exception as e:
    #         logging.warning("End date change error: %s", e)

    # start_date_entry.bind('<<DateEntrySelected>>', on_start_date_change)
    # end_date_entry.bind('<<DateEntrySelected>>', on_end_date_change)

    from datetime import datetime, timedelta
    today = datetime.now()
    set_start_date(today.strftime("%Y/%m/%d"))
    set_end_date(today.strftime("%Y/%m/%d"))

    btns = ttk.Frame(date_frame)
    btns.grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(0, 8))

    def set_week():
        d = datetime.now()
        s = d - timedelta(days=d.weekday())
        e = s + timedelta(days=6)
        set_start_date(s.strftime("%Y/%m/%d"))
        set_end_date(e.strftime("%Y/%m/%d"))

    def set_month():
        d = datetime.now()
        s = d.replace(day=1)
        n = s.replace(year=s.year + 1, month=1) if s.month == 12 else s.replace(month=s.month + 1)
        e = n - timedelta(days=1)
        set_start_date(s.strftime("%Y/%m/%d"))
        set_end_date(e.strftime("%Y/%m/%d"))

    def set_last_days(days):
        d = datetime.now()
        s = d - timedelta(days=days - 1)
        set_start_date(s.strftime("%Y/%m/%d"))
        set_end_date(d.strftime("%Y/%m/%d"))

    make_button(btns, text="本周", command=set_week, width=10, role="primary").pack(side=tk.LEFT, padx=(0, 8))
    make_button(btns, text="本月", command=set_month, width=10, role="primary").pack(side=tk.LEFT, padx=(0, 8))
    make_button(btns, text="近30天", command=lambda: set_last_days(30), width=10, role="primary").pack(side=tk.LEFT, padx=(0, 8))
    make_button(btns, text="近90天", command=lambda: set_last_days(90), width=10, role="primary").pack(side=tk.LEFT, padx=(0, 8))
    make_button(btns, text="恢复默认", command=lambda: (set_start_date(today.strftime("%Y/%m/%d")),
                                                        set_end_date(today.strftime("%Y/%m/%d"))),
                                                       width=10, role="primary").pack(side=tk.LEFT)

    quick_map = ttk.LabelFrame(main_tab, text="▌快速产品线映射", padding=16)
    quick_map.pack(fill=tk.X, pady=(0, 25))
    quick_map.columnconfigure(1, weight=1)
    quick_map.columnconfigure(3, weight=1)
    product_quick_var = tk.StringVar()
    contact_quick_var = tk.StringVar()
    ttk.Label(quick_map, text="产品线名称:").grid(row=0, column=0, sticky=tk.W, pady=(8, 10))
    ttk.Entry(quick_map, textvariable=product_quick_var, width=30, font=ENTRY_FONT).grid(
        row=0, column=1, sticky=tk.EW, padx=(8, 8), pady=(8, 10))
    ttk.Label(quick_map, text="新对接人:").grid(row=0, column=2, sticky=tk.W, pady=(8, 10))
    ttk.Entry(quick_map, textvariable=contact_quick_var, width=30, font=ENTRY_FONT).grid(
        row=0, column=3, sticky=tk.EW, padx=(8, 8), pady=(8, 10))

    def add_quick_mapping():
        p = product_quick_var.get().strip()
        c = contact_quick_var.get().strip()
        if not p or not c:
            messagebox.showerror("错误", "请填写产品线与新对接人")
            return
        pl_manager.add_row(p, c)
        product_quick_var.set("")
        contact_quick_var.set("")
        _create_toast(root, "已添加到映射列表")

    make_button(quick_map, text="添加到映射列表", command=add_quick_mapping, width=16, role="primary").grid(
        row=0, column=4, padx=(8, 0))

    mapping_frame = ttk.LabelFrame(mapping_tab, text="▌多产品线输入组件", padding=16)
    mapping_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 25))
    pl_manager = ProductLineManager(mapping_frame)
    pl_manager.add_row()
    pl_manager.load_from_file()
    mapping_count_var = tk.StringVar(value=f"当前映射条数：{len(pl_manager.rows)}")
    ttk.Label(mapping_frame, textvariable=mapping_count_var).pack(anchor=tk.W, pady=(4, 8))

    def update_mapping_count():
        try:
            mapping_count_var.set(f"当前映射条数：{len(pl_manager.rows)}")
        except Exception:
            pass

    mf_actions = ttk.Frame(mapping_frame)
    mf_actions.pack(fill=tk.X, pady=(12, 0))

    def add_row_action():
        pl_manager.add_row()
        update_mapping_count()
        _create_toast(root, "已新增一行")

    def clear_all_action():
        pl_manager.clear()
        update_mapping_count()
        _create_toast(root, "已清空映射")

    def save_action():
        try:
            pl_manager.save_to_file()
            _create_toast(root, "保存成功")
        except Exception:
            _create_toast(root, "保存失败")

    def load_action():
        pl_manager.clear()
        pl_manager.load_from_file()
        update_mapping_count()
        _create_toast(root, f"已从文件加载{len(pl_manager.rows)}条")

    ttk.Button(mf_actions, text="新增一行", command=add_row_action, width=12).pack(side=tk.LEFT)
    ttk.Button(mf_actions, text="清空全部", command=clear_all_action, width=12).pack(side=tk.LEFT, padx=(8, 0))
    ttk.Button(mf_actions, text="保存到文件", command=save_action, width=12).pack(side=tk.LEFT, padx=(8, 0))
    ttk.Button(mf_actions, text="从文件加载", command=load_action, width=12).pack(side=tk.LEFT, padx=(8, 0))

    ttk.Label(mapping_frame, text="替换模式:").pack(anchor=tk.W, pady=(10, 4))
    mode_bar = ttk.Frame(mapping_frame)
    mode_bar.pack(fill=tk.X, pady=(0, 8))
    ttk.Radiobutton(mode_bar, text="覆盖所有", variable=replace_mode_var,
                     value='overwrite', command=save_app_state).pack(side=tk.LEFT)
    ttk.Radiobutton(mode_bar, text="仅填空值", variable=replace_mode_var,
                     value='fill_empty', command=save_app_state).pack(side=tk.LEFT, padx=(12, 0))

    def preview_mappings():
        try:
            inp = app_state.input_file or input_entry.get().strip()
            if not inp or not os.path.exists(inp):
                messagebox.showerror("错误", "请先选择有效的输入文件")
                return
            sheets = get_sheets_with_data(inp)
            if not sheets:
                messagebox.showerror("错误", "未找到包含数据的工作表")
                return
            import pandas as pd
            from core import transform as transform_core
            column_mapper_local = ColumnMapper()
            total = 0
            affected = {p: 0 for p, _ in pl_manager.get_mappings()}
            for s in sheets:
                try:
                    df = pd.read_excel(inp, sheet_name=s, header=1, converters={'发起时间': str})
                    df = transform_core.deep_clean_columns(df)
                    matched = transform_core.dynamic_column_matching(df, column_mapper_local)
                    src_col = matched.get('产品线')
                    if not src_col:
                        for col in df.columns:
                            if '产品' in str(col):
                                src_col = col
                                break
                    if not src_col:
                        continue
                    total += len(df)
                    ser = df[src_col].astype(str).str.strip().str.lower()
                    for p, _c in pl_manager.get_mappings():
                        cnt = (ser == str(p).strip().lower()).sum()
                        affected[p] = affected.get(p, 0) + int(cnt)
                except Exception:
                    continue
            lines = [f"总行数：{total}"] + [f"{p}：{n} 行" for p, n in affected.items()]
            messagebox.showinfo("预览", "\n".join(lines))
        except Exception as e:
            messagebox.showerror("错误", str(e))

    make_button(mapping_frame, text="映射预览", command=preview_mappings, width=12, role="info").pack(anchor=tk.W)

    column_mapper = ColumnMapper()
    _build_column_mapping_tab(column_tab, column_mapper)
    _build_info_tab(info_tab)

    overlay = tk.Frame(root, bg=BG_COLOR, padx=16, pady=16)
    overlay.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor="se")
    exit_btn = make_button(overlay, text="退出", command=root.quit, width=14, role="danger")
    process_btn = make_button(overlay, text="开始处理", command=lambda: start_process(), width=18, role="primary")
    process_btn.grid(row=0, column=1, padx=(25, 0), pady=(8, 8))
    exit_btn.grid(row=0, column=0, padx=(25, 0), pady=(8, 8))

    processor = ExcelProcessor(process_raw_excel)

    def start_process():
        inp = app_state.input_file or input_entry.get().strip()
        outp = app_state.output_dir or output_entry.get().strip()
        if not inp or not outp:
            messagebox.showerror("错误", "请选择输入文件和输出目录！")
            return
        if not os.path.exists(inp):
            messagebox.showerror("错误", "输入文件不存在！")
            return
        if not os.path.exists(outp):
            messagebox.showerror("错误", "输出目录不存在！")
            return
        process_btn.configure(state="disabled")
        exit_btn.configure(state="disabled")

        try:
            ov, upd, cancel_event = _show_processing_overlay(root, process_btn, exit_btn, None)
        except Exception as e:
            messagebox.showerror("错误", f"创建进度窗口失败: {e}")
            process_btn.configure(state="normal")
            exit_btn.configure(state="normal")
            return

        try:
            def get_date_value(widget):
                try:
                    if hasattr(widget, 'get'):
                        val = widget.get()
                        if val:
                            return val
                    if hasattr(widget, 'entry') and hasattr(widget.entry, 'get'):
                        val = widget.entry.get()
                        if val:
                            return val
                except Exception:
                    pass
                return None

            sd = get_date_value(start_date_entry) or start_date_var.get() or app_state.start_date or today.strftime("%Y/%m/%d")
            ed = get_date_value(end_date_entry) or end_date_var.get() or app_state.end_date or today.strftime("%Y/%m/%d")
        except Exception:
            sd = start_date_var.get() or app_state.start_date or today.strftime("%Y/%m/%d")
            ed = end_date_var.get() or app_state.end_date or today.strftime("%Y/%m/%d")
        sdt = datetime.strptime(sd, "%Y/%m/%d")
        edt = datetime.strptime(ed, "%Y/%m/%d")
        out_file = f"{outp}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
        ok = False

        valid, msg = processor.validate_mappings(pl_manager.get_mappings())
        if not valid:
            messagebox.showerror("错误", msg)
            try:
                ov.destroy()
            except Exception:
                pass
            process_btn.configure(state="normal")
            exit_btn.configure(state="normal")
            return

        try:
            ok = processor.process(
                inp,
                out_file,
                sdt,
                edt,
                pl_manager.get_mappings(),
                replace_mode=replace_mode_var.get(),
                progress_callback=upd,
                cancel_event=cancel_event,
            )
        except Exception as e:
            messagebox.showerror("错误", str(e))
        finally:
            try:
                ov.destroy()
            except Exception:
                pass
            process_btn.configure(state="normal")
            exit_btn.configure(state="normal")

        if ok:
            try:
                ans = messagebox.askyesno("完成", "处理完成，是否立即打开文件？")
                if ans:
                    os.startfile(out_file)
                ans2 = messagebox.askyesno("完成", "是否打开输出目录？")
                if ans2:
                    try:
                        os.startfile(os.path.dirname(out_file))
                    except Exception:
                        pass
            except Exception:
                pass
        else:
            _create_toast(root, "处理已取消或失败", 2500)

    load_app_state()
    root.bind("<Return>", lambda e: start_process())
    root.bind("<Escape>", lambda e: root.quit())
    return root


if __name__ == "__main__":
    app = create_app()
    app.mainloop()
