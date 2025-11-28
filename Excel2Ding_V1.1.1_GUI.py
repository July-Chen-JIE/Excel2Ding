import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import re
import traceback
from tkinter import ttk
from tkcalendar import DateEntry
import os

# 列映射配置
COLUMN_MAPPING = {
    '发起人姓名': '对接人',
    '发起时间': '创建时间',
    '项目名称': '项目名称',
    '产品线': '产品',
    '建议报价元': '报价金额',
    '申请状态': '当前进度'
}

def deep_clean_columns(df):
    """深度清洗列名"""
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
    
    for target, aliases in column_alias.items():
        found = False
        for col in df.columns:
            col_clean = re.sub(r'[\s：()（）\n\t]', '', str(col)).strip()
            for alias in aliases:
                alias_clean = re.sub(r'[\s：()（）\n\t]', '', str(alias)).strip()
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
    """处理Excel文件的主函数"""
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
        
        # 时间范围过滤
        start_dt = datetime.strptime(start_date, "%Y/%m/%d")
        end_dt = datetime.strptime(end_date, "%Y/%m/%d")
        mask = (valid_df['datetime_obj'].dt.date >= start_dt.date()) & \
               (valid_df['datetime_obj'].dt.date <= end_dt.date())
        filtered = valid_df[mask].copy()  # 创建副本避免警告
        
        # 生成输出数据 (85%)
        if progress_callback:
            progress_callback(85, "正在生成输出数据...")
        filtered.loc[:, '当前周'] = filtered['datetime_obj'].dt.isocalendar().week
        output_df = filtered[list(matched.values())].rename(columns=matched)
        output_df = output_df.rename(columns=COLUMN_MAPPING)
        output_df.insert(2, '当前周', filtered['当前周'])
        output_df['创建时间'] = filtered['datetime_obj'].dt.strftime('%Y/%m/%d %H:%M')
        
        # 按目标列排序
        final_columns = ['对接人', '创建时间', '当前周', '项目名称', '产品', '当前进度', '报价金额']
        output_df = output_df[final_columns]
        
        # 保存文件 (95%)
        if progress_callback:
            progress_callback(95, "正在保存文件...")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            
            # 设置格式
            from openpyxl.styles import Alignment
            alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # 自动调整列宽
            for idx, col in enumerate(output_df.columns):
                max_length = 0
                column = chr(65 + idx)
                
                # 计算最大列宽
                max_length = max(
                    max_length,
                    len(str(col)) * 2,
                    max(len(str(cell)) * 1.2 for cell in output_df[col].astype(str))
                )
                
                # 设置列宽（限制最大宽度为50）
                adjusted_width = min(max_length + 4, 50)
                worksheet.column_dimensions[column].width = adjusted_width
                
                # 设置对齐方式
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
        messagebox.showerror("错误", f"处理失败: {str(e)}")
        return False

def create_gui():
    """创建现代化GUI界面"""
    root = tk.Tk()
    
    # 定义日期操作函数
    def set_week_start_end():
        """设置本周起止日期"""
        today = datetime.now()
        week_start = today - timedelta(days=today.weekday())
        week_end = week_start + timedelta(days=6)
        start_date.set_date(week_start)
        end_date.set_date(week_end)
    
    def set_month_start_end():
        """设置本月起止日期"""
        today = datetime.now()
        month_start = today.replace(day=1)
        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)
        month_end = next_month - timedelta(days=1)
        start_date.set_date(month_start)
        end_date.set_date(month_end)
    
    def clear_dates():
        """清除日期选择"""
        today = datetime.now()
        start_date.set_date(today)
        end_date.set_date(today)
    
    def update_progress(progress_var, progress_label, value, message):
        """更新进度条和提示文本"""
        progress_var.set(value)
        progress_label.configure(text=f"⏳ {message}")
        root.update()
    
    def start_process():
        """开始处理函数"""
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
        
        # 禁用按钮
        process_btn.configure(state='disabled')
        exit_btn.configure(state='disabled')
        
        # 重置进度条
        progress_var.set(0)
        progress_label.configure(text="⏳ 正在处理...")
        root.update()
        
        try:
            if process_excel(
                input_file,
                start_date.get(),
                end_date.get(),
                f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                progress_callback=lambda p, msg: update_progress(progress_var, progress_label, p, msg)
            ):
                messagebox.showinfo("完成", "文件处理成功！")
        finally:
            # 恢复按钮状态
            process_btn.configure(state='normal')
            exit_btn.configure(state='normal')
            # 重置进度条文本
            progress_label.configure(text="⏳ 准备就绪")

    # 配置现代感配色方案
    PRIMARY_COLOR = "#409EFF"      # 主色调（科技蓝）
    SECONDARY_COLOR = "#6DD3B7"    # 辅助色（清新绿）
    BG_COLOR = "#F5F7FA"           # 背景色（浅灰）
    TEXT_COLOR = "#2D3748"         # 主文本色
    BUTTON_TEXT_COLOR = "white"    # 按钮文字颜色
    
    # 设置窗口样式
    root.title("定制审批单处理工具 v2.0 Design by czj")
    root.geometry("600x800")
    root.configure(bg=BG_COLOR)
    
    # 配置全局样式
    style = ttk.Style()
    style.theme_use('clam')
    
    # 设置控件样式
    style.configure('TFrame', background=BG_COLOR)
    style.configure('TLabel', 
                   font=('Segoe UI', 11),
                   padding=5,
                   background=BG_COLOR)
    style.configure('TLabelframe', 
                   background=BG_COLOR,
                   padding=15)
    style.configure('TLabelframe.Label',
                   font=('Segoe UI', 12, 'bold'),
                   foreground=PRIMARY_COLOR,
                   background=BG_COLOR)
    style.configure('Modern.TButton',
                   font=('Segoe UI', 11, 'bold'),
                   padding=10,
                   background=PRIMARY_COLOR,
                   foreground=BUTTON_TEXT_COLOR,  # 设置按钮文字颜色为白色
                   borderwidth=0,  # 移除边框
                   relief="flat")  # 扁平化效果
    
    # 添加按钮圆角效果（通过自定义布局）
    style.layout('Modern.TButton', [
        ('Button.padding', {'children': [
            ('Button.label', {'sticky': 'nswe'})
        ], 'sticky': 'nswe'})])
    
    # 配置按钮鼠标悬停效果
    style.map('Modern.TButton',
              background=[('active', '#66B1FF'), ('pressed', '#3a8ee6')],
              foreground=[('active', 'white'), ('pressed', 'white')])
    
    # 如果要对进度条也添加圆角效果，可以添加以下配置
    style.configure('Modern.Horizontal.TProgressbar',
                   troughcolor='#E2E8F0',
                   background=SECONDARY_COLOR,
                   thickness=25,
                   borderwidth=0,
                   relief="flat")
    
    # 创建主框架
    main_frame = ttk.Frame(root, padding=25)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 输入文件部分
    input_frame = ttk.LabelFrame(main_frame, text="▌输入文件设置")
    input_frame.pack(fill=tk.X, pady=10)
    
    ttk.Label(input_frame, text="📂 输入文件:").pack(anchor="w", pady=(5, 2))
    input_entry = ttk.Entry(input_frame, width=50, font=('Segoe UI', 10))
    input_entry.pack(side=tk.LEFT, pady=5, padx=(0, 5), fill=tk.X, expand=True)
    
    ttk.Button(input_frame, 
              text="浏览",
              command=lambda: (input_entry.delete(0, tk.END),
                             input_entry.insert(0, filedialog.askopenfilename(
                                 filetypes=[("Excel文件", "*.xlsx")]))),
              style='Modern.TButton').pack(side=tk.RIGHT, padx=5)
    
    # 日期选择部分
    date_frame = ttk.LabelFrame(main_frame, text="▌日期范围设置")
    date_frame.pack(fill=tk.X, pady=10)
    
    # 日期选择器样式
    date_style = {
        'font': ('Segoe UI', 10),
        'background': 'white',
        'foreground': TEXT_COLOR,
        'selectbackground': PRIMARY_COLOR,
        'date_pattern': 'yyyy/mm/dd'
    }
    
    # 日期选择器
    date_select_frame = ttk.Frame(date_frame)
    date_select_frame.pack(fill=tk.X, pady=5)
    
    # 开始日期
    ttk.Label(date_select_frame, text="📅 开始日期:").pack(side=tk.LEFT)
    start_date = DateEntry(date_select_frame, **date_style)
    start_date.pack(side=tk.LEFT, padx=(5, 20))
    
    # 结束日期
    ttk.Label(date_select_frame, text="📅 结束日期:").pack(side=tk.LEFT)
    end_date = DateEntry(date_select_frame, **date_style)
    end_date.pack(side=tk.LEFT, padx=5)
    
    # 日期快捷按钮
    date_buttons_frame = ttk.Frame(date_frame)
    date_buttons_frame.pack(fill=tk.X, pady=5)
    
    ttk.Button(date_buttons_frame,
              text="📅 本周",
              command=set_week_start_end,
              style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 5))
    
    ttk.Button(date_buttons_frame,
              text="📆 本月",
              command=set_month_start_end,
              style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 5))
    
    ttk.Button(date_buttons_frame,
              text="🔄 恢复",
              command=clear_dates,
              style='Modern.TButton').pack(side=tk.LEFT)
    
    # 输出设置部分
    output_frame = ttk.LabelFrame(main_frame, text="▌输出设置")
    output_frame.pack(fill=tk.X, pady=10)
    
    ttk.Label(output_frame, text="💾 保存路径:").pack(anchor="w", pady=(5, 2))
    output_entry = ttk.Entry(output_frame, width=50, font=('Segoe UI', 10))
    output_entry.pack(side=tk.LEFT, pady=5, padx=(0, 5), fill=tk.X, expand=True)
    
    ttk.Button(output_frame,
              text="浏览",
              command=lambda: (output_entry.delete(0, tk.END),
                             output_entry.insert(0, filedialog.askdirectory())),
              style='Modern.TButton').pack(side=tk.RIGHT, padx=5)
    
    # 进度条部分
    progress_frame = ttk.LabelFrame(main_frame, text="▌处理进度")
    progress_frame.pack(fill=tk.X, pady=10)
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        progress_frame,
        variable=progress_var,
        maximum=100,
        mode='determinate',
        style='Modern.Horizontal.TProgressbar'
    )
    progress_bar.pack(fill=tk.X, pady=(5, 0))
    
    progress_label = ttk.Label(
        progress_frame,
        text="⏳ 准备就绪"
    )
    progress_label.pack(anchor="w", pady=(5, 0))
    
    # 操作按钮部分
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=15)
    
    process_btn = ttk.Button(
        button_frame,
        text="🚀 开始处理",
        command=start_process,
        style='Modern.TButton'
    )
    process_btn.pack(side=tk.RIGHT, padx=8)
    
    exit_btn = ttk.Button(
        button_frame,
        text="x 退出程序",
        command=root.quit,
        style='Modern.TButton'
    )
    exit_btn.pack(side=tk.RIGHT)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()
