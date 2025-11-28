import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import re
import traceback
from tkinter import ttk
from tkcalendar import DateEntry
import os

COLUMN_MAPPING = {
    '发起人姓名': '对接人',
    '发起时间': '创建时间',
    '项目名称': '项目名称',
    '产品线': '产品',
    '建议报价元': '报价金额',
    '申请状态': '当前进度'
}

REQUIRED_COLUMNS = list(COLUMN_MAPPING.keys())

def deep_clean_columns(df):
    """深度清洗列名"""
    # 清理列名中的特殊字符
    df.columns = [re.sub(r'[\s：()（）\n\t]', '', str(col)).strip() for col in df.columns]
    return df.dropna(how='all')

def dynamic_column_matching(df):
    """精确列名匹配"""
    column_alias = {
        '发起人姓名': ['发起人姓名'],
        '发起时间': ['发起时间'],
        '项目名称': ['项目名称'],
        '产品线': ['产品线'],
        '建议报价元': ['建议报价(元)'],
        '申请状态': ['申请状态']
    }
    matched = {}
    print("输入文件的列名：", df.columns.tolist())
    
    # 修改匹配逻辑以处理更复杂的列名
    for target, aliases in column_alias.items():
        found = False
        for col in df.columns:
            col_clean = re.sub(r'[\s：()（）\n\t]', '', str(col)).strip()  # 添加 \t
            for alias in aliases:
                alias_clean = re.sub(r'[\s：()（）\n\t]', '', str(alias)).strip()  # 添加 \t
                if col_clean == alias_clean:
                    matched[target] = col
                    found = True
                    break
            if found:
                break
        if not found:
            raise ValueError(f"列[{target}]未找到，当前列：{df.columns.tolist()}")
    
    print("匹配结果：", matched)
    return matched

def excel_serial_to_datetime(serial):
    """将 Excel 序列号转换为 datetime 对象"""
    try:
        if isinstance(serial, str):
            return pd.to_datetime(serial)
        
        base_date = datetime(1899, 12, 30)
        if pd.isna(serial):
            return pd.NaT
            
        days = int(serial)
        fractional_day = serial - days
        hours = int(fractional_day * 24)
        minutes = int((fractional_day * 24 - hours) * 60)
        seconds = int(((fractional_day * 24 - hours) * 60 - minutes) * 60)
        return base_date + timedelta(days=days, hours=hours, minutes=minutes, seconds=seconds)
    except Exception as e:
        print(f"警告：序列号 {serial} 转换失败：{str(e)}")
        return pd.NaT

def process_excel(input_path, start_date, end_date, output_path, progress_callback=None):
    try:
        # 读取文件 (10%)
        if progress_callback:
            progress_callback(10, "正在读取文件...")
        df = pd.read_excel(input_path, engine='openpyxl')
        df = df.dropna(how='all', axis=1)
        
        # 清洗列名 (20%)
        if progress_callback:
            progress_callback(20, "正在清洗数据...")
        df = deep_clean_columns(df)
        
        # 列匹配 (30%)
        if progress_callback:
            progress_callback(30, "正在匹配列名...")
        matched = dynamic_column_matching(df)
        
        # 日期处理 (50%)
        if progress_callback:
            progress_callback(50, "正在处理日期...")
        try:
            df['datetime_obj'] = df[matched['发起时间']].apply(
                lambda x: excel_serial_to_datetime(float(x)) if pd.notna(x) else pd.NaT
            )
        except Exception as e:
            print("日期解析失败：", e)
            raise ValueError("日期列格式不正确，请检查输入文件的日期格式！")
        
        # 数据过滤和转换 (70%)
        if progress_callback:
            progress_callback(70, "正在过滤数据...")
        valid_df = df[df['datetime_obj'].notna()]
        if valid_df.empty:
            raise ValueError("日期解析失败，请检查“发起时间”列是否为有效的 Excel 序列号格式")
        
        # 修改时间范围过滤部分的日期格式解析
        start_dt = datetime.strptime(start_date, "%Y/%m/%d")  # 修改为与 DateEntry 相同的格式
        end_dt = datetime.strptime(end_date, "%Y/%m/%d")      # 修改为与 DateEntry 相同的格式
        mask = (valid_df['datetime_obj'].dt.date >= start_dt.date()) & \
               (valid_df['datetime_obj'].dt.date <= end_dt.date())
        filtered = valid_df[mask]
        
        # 生成输出数据 (85%)
        if progress_callback:
            progress_callback(85, "正在生成输出数据...")
        filtered['当前周'] = filtered['datetime_obj'].dt.isocalendar().week
        output_df = filtered[list(matched.values())].rename(columns=matched)
        output_df = output_df.rename(columns=COLUMN_MAPPING)
        output_df.insert(2, '当前周', filtered['当前周'])
        output_df['创建时间'] = filtered['datetime_obj'].dt.strftime('%Y/%m/%d %H:%M')
        
        # 按目标列排序
        final_columns = ['对接人', '创建时间', '当前周', '项目名称', '产品', '当前进度', '报价金额']
        missing_columns = [col for col in final_columns if col not in output_df.columns]
        if missing_columns:
            print("输出数据的列名：", output_df.columns.tolist())  # 调试信息
            raise ValueError(f"以下列未找到：{missing_columns}")
        
        output_df = output_df[final_columns]
        
        # 保存文件 (95%)
        if progress_callback:
            progress_callback(95, "正在保存文件...")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 保存数据
            output_df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            
            # 设置基本对齐方式
            from openpyxl.styles import Alignment
            alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # 自动调整列宽
            for idx, col in enumerate(output_df.columns):
                max_length = 0
                column = chr(65 + idx)  # 转换为 Excel 列字母 (A, B, C, ...)
                
                # 获取列标题的长度（考虑中文字符）
                max_length = max(max_length, len(str(col)) * 2)
                
                # 获取该列所有数据的最大长度
                for row in output_df[col].astype(str):
                    # 中文字符计数为2，其他字符计数为1
                    length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in str(row))
                    max_length = max(max_length, length)
                
                # 设置列宽（根据字符长度适当调整）
                adjusted_width = min(max_length + 4, 50)  # 限制最大宽度为50
                worksheet.column_dimensions[column].width = adjusted_width
                
                # 设置单元格对齐方式
                for cell in worksheet[column]:
                    cell.alignment = alignment
        
        # 完成 (100%)
        if progress_callback:
            progress_callback(100, "处理完成！")
        
        return True
        
    except Exception as e:
        if progress_callback:
            progress_callback(0, f"处理失败: {str(e)}")
        traceback.print_exc()
        error_msg = f"""处理失败: {str(e)}
        常见问题:
        1. 确认Excel前1行为非数据行
        2. 检查列名是否包含: 发起人姓名,发起时间,项目名称,产品线,建议报价元,申请状态
        3. 日期格式必须为 Excel 序列号（如 45757.6148148148）"""
        messagebox.showerror("错误", error_msg)
        return False

def create_gui():
    """创建符合 Element 设计规范的 GUI 界面"""
    root = tk.Tk()
    
    def set_week_start_end():
        """设置本周起止日期"""
        today = datetime.now()
        # 获取本周一
        week_start = today - timedelta(days=today.weekday())
        # 获取本周日
        week_end = week_start + timedelta(days=6)
        
        start_date.set_date(week_start)
        end_date.set_date(week_end)
    
    def set_month_start_end():
        """设置本月起止日期"""
        today = datetime.now()
        # 获取本月第一天
        month_start = today.replace(day=1)
        # 获取下月第一天
        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)
        # 获取本月最后一天
        month_end = next_month - timedelta(days=1)
        
        start_date.set_date(month_start)
        end_date.set_date(month_end)
        
    def clear_dates():
        """清除日期选择"""
        today = datetime.now()
        start_date.set_date(today)
        end_date.set_date(today)
        
    root.title("定制审批单处理工具 v2.6")
    root.geometry("700x650")  # 略微增加窗口大小
    root.configure(bg="#FFFFFF")  # 使用 Element 的白色背景
    
    # 设置统一的样式
    style = ttk.Style()
    style.configure("Element.TLabel",
                   font=("Microsoft YaHei", 12),
                   background="#FFFFFF",
                   foreground="#303133")  # Element 主文本颜色
    
    style.configure("Element.TButton",
                   font=("Microsoft YaHei", 11),
                   padding=6)
    
    # 添加进度条样式
    style.configure(
        "Element.Horizontal.TProgressbar",
        troughcolor="#F5F7FA",  # 进度条背景色
        background="#409EFF",   # 进度条前景色
        thickness=20            # 进度条高度
    )
    
    # 创建主框架
    main_frame = ttk.Frame(root, padding="20 20 20 20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 输入文件部分
    input_frame = ttk.LabelFrame(main_frame, text="输入设置", padding="10 10 10 10")
    input_frame.pack(fill=tk.X, pady=(0, 15))
    
    ttk.Label(input_frame, text="输入文件:", style="Element.TLabel").pack(anchor="w", pady=(5, 2))
    input_entry = ttk.Entry(input_frame, width=50, font=("Microsoft YaHei", 10))
    input_entry.pack(side=tk.LEFT, pady=5, padx=(0, 5), fill=tk.X, expand=True)
    # 修改输入文件浏览按钮
    ttk.Button(input_frame, text="浏览", 
               command=lambda: (input_entry.delete(0, tk.END), 
                              input_entry.insert(0, filedialog.askopenfilename(
                                  filetypes=[("Excel文件", "*.xlsx")]))),
               style="Element.TButton").pack(side=tk.RIGHT, pady=5)
    
    # 日期选择部分
    date_frame = ttk.LabelFrame(main_frame, text="日期范围", padding="10 10 10 10")
    date_frame.pack(fill=tk.X, pady=(0, 15))
    
    # 设置日期选择器样式
    date_style = {"font": ("Microsoft YaHei", 10),
                  "background": "white",
                  "foreground": "#303133",
                  "selectbackground": "#409EFF",  # Element 主题蓝
                  "date_pattern": "yyyy/mm/dd"}  # 修改日期格式为 YYYY/MM/DD
    
    # 添加日期操作按钮框架
    date_buttons_frame = ttk.Frame(date_frame)
    date_buttons_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))  # 改为底部布局

    # 添加日期操作按钮
    ttk.Button(date_buttons_frame,
              text="本周",
              command=set_week_start_end,
              style="Element.TButton").pack(side=tk.LEFT, padx=(0, 5))  # 改为左对齐

    ttk.Button(date_buttons_frame,
              text="本月",
              command=set_month_start_end,
              style="Element.TButton").pack(side=tk.LEFT, padx=(0, 5))  # 改为左对齐

    ttk.Button(date_buttons_frame,
              text="清除",
              command=clear_dates,
              style="Element.TButton").pack(side=tk.LEFT)  # 改为左对齐
    
    # 日期选择器左侧框架
    date_left_frame = ttk.Frame(date_frame)
    date_left_frame.pack(side=tk.LEFT, padx=(0, 10))
    ttk.Label(date_left_frame, text="开始日期:", style="Element.TLabel").pack(anchor="w")
    start_date = DateEntry(date_left_frame, **date_style)
    start_date.pack(pady=(5, 0))
    
    # 日期选择器右侧框架
    date_right_frame = ttk.Frame(date_frame)
    date_right_frame.pack(side=tk.LEFT)
    ttk.Label(date_right_frame, text="结束日期:", style="Element.TLabel").pack(anchor="w")
    end_date = DateEntry(date_right_frame, **date_style)
    end_date.pack(pady=(5, 0))
    
    # 输出设置部分
    output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="10 10 10 10")
    output_frame.pack(fill=tk.X, pady=(0, 15))
    
    ttk.Label(output_frame, text="保存路径:", style="Element.TLabel").pack(anchor="w", pady=(5, 2))
    output_entry = ttk.Entry(output_frame, width=50, font=("Microsoft YaHei", 10))
    output_entry.pack(side=tk.LEFT, pady=5, padx=(0, 5), fill=tk.X, expand=True)
    # 修改输出目录浏览按钮
    ttk.Button(output_frame, text="浏览",
               command=lambda: (output_entry.delete(0, tk.END),
                              output_entry.insert(0, filedialog.askdirectory())),
               style="Element.TButton").pack(side=tk.RIGHT, pady=5)
    
    # 在操作按钮部分之前添加进度条框架
    progress_frame = ttk.LabelFrame(main_frame, text="处理进度", padding="10 10 10 10")
    progress_frame.pack(fill=tk.X, pady=(0, 15))
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        progress_frame,
        variable=progress_var,
        maximum=100,
        mode='determinate',
        style='Element.Horizontal.TProgressbar'
    )
    progress_bar.pack(fill=tk.X, pady=(5, 0))
    
    # 添加进度标签
    progress_label = ttk.Label(
        progress_frame,
        text="准备就绪",
        style="Element.TLabel"
    )
    progress_label.pack(anchor="w", pady=(5, 0))
    
    # 操作按钮部分
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))
    
    def start_process():
        input_file = input_entry.get().strip()
        output_dir = output_entry.get().strip()
        
        # 验证路径
        if not input_file or not output_dir:
            messagebox.showerror("错误", "请选择输入文件和输出目录！")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "输入文件不存在！")
            return
        
        if not os.path.exists(output_dir):
            messagebox.showerror("错误", "输出目录不存在！")
            return
        
        # 禁用按钮，防止重复点击
        process_btn.configure(state='disabled')
        exit_btn.configure(state='disabled')
        
        # 重置进度条
        progress_var.set(0)
        progress_label.configure(text="正在处理...")
        root.update()
        
        try:
            if process_excel(
                input_file,
                start_date.get(),
                end_date.get(),
                f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                progress_callback=lambda p, msg: update_progress(progress_var, progress_label, p, msg)
            ):
                messagebox.showinfo("处理完成", "文件处理成功！")
        finally:
            # 恢复按钮状态
            process_btn.configure(state='normal')
            exit_btn.configure(state='normal')
            # 完成后重置进度条和文本
            progress_var.set(0)
            progress_label.configure(text="准备就绪")
            root.update()
    
    # 添加进度更新函数
    def update_progress(progress_var, progress_label, value, message):
        progress_var.set(value)
        progress_label.configure(text=message)
        root.update()
    
    # 使用 Element 的主题色
    process_btn = tk.Button(button_frame,
                          text="开始处理",
                          command=start_process,
                          bg="#409EFF",
                          fg="white",
                          font=("Microsoft YaHei", 11),
                          relief=tk.FLAT,
                          padx=20,
                          pady=8)
    process_btn.pack(side=tk.RIGHT, padx=5)
    
    exit_btn = tk.Button(button_frame,
                        text="退出",
                        command=root.quit,
                        bg="#F56C6C",  # Element 的红色
                        fg="white",
                        font=("Microsoft YaHei", 11),
                        relief=tk.FLAT,
                        padx=20,
                        pady=8)
    exit_btn.pack(side=tk.RIGHT, padx=5)
    
    # 添加鼠标悬停效果
    def on_enter(e):
        e.widget['bg'] = '#66B1FF' if e.widget['bg'] == '#409EFF' else '#F78989'
    
    def on_leave(e):
        e.widget['bg'] = '#409EFF' if e.widget['text'] == '开始处理' else '#F56C6C'
    
    for btn in [process_btn, exit_btn]:
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()
