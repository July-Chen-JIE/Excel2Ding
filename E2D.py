#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据处理工具 GUI版本（柔和简约风）
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime, timedelta
import warnings
import re
import traceback
from tkinter import ttk
from tkcalendar import DateEntry
import json
from openpyxl.styles import Alignment
from openpyxl import load_workbook


warnings.filterwarnings('ignore')

# -------------------------- 柔和简约风设计规范 --------------------------
# 色彩：低饱和莫兰迪色系，避免刺眼对比
PRIMARY_COLOR = "#4285F4"       # 柔和主蓝（Google蓝，低饱和不刺眼）
PRIMARY_LIGHT = "#73A1FF"      # 主色浅态（hover）
PRIMARY_DARK = "#3367D6"       # 主色深态（点击）
SUCCESS_COLOR = "#34A853"       # 柔和成功绿
SUCCESS_LIGHT = "#5BC86D"      # 成功绿浅态
SUCCESS_DARK = "#2A8644"       # 成功绿深态
DANGER_COLOR = "#EA4335"       # 柔和危险红
DANGER_LIGHT = "#F1695E"      # 危险红浅态
DANGER_DARK = "#D33526"        # 危险红深态
TEXT_PRIMARY = "#2D3748"       # 主要文本（深灰，不刺眼）
TEXT_REGULAR = "#718096"       # 常规文本（中灰）
TEXT_SECONDARY = "#A0AEC0"     # 次要文本（浅灰）
BG_MAIN = "#FAFAFA"            # 主背景（极浅灰，比纯白柔和）
BG_CARD = "#FFFFFF"            # 卡片背景（纯白，突出内容）
BG_INPUT = "#F7FAFC"           # 输入框背景（极浅灰）
BORDER_BASE = "#E2E8F0"        # 基础边框（极浅灰，弱化边框感）

# 尺寸：宽松呼吸感，避免拥挤
PADDING_XS = 6
PADDING_SM = 12
PADDING_MD = 20
PADDING_LG = 28
BORDER_RADIUS = 8              # 统一圆角（现代感）
SHADOW_SOFT = "0 2px 10px rgba(0,0,0,0.05)"  # 超柔和阴影（不突兀）

# 字体：轻盈现代
FONT_FAMILY = "Microsoft YaHei UI, PingFang SC, Roboto, sans-serif"
TITLE_FONT = (FONT_FAMILY, 19, 'bold')       # 页面标题（轻盈醒目）
SUBTITLE_FONT = (FONT_FAMILY, 14, 'bold')    # 卡片标题
CONTENT_FONT = (FONT_FAMILY, 13)             # 内容文本
BUTTON_FONT = (FONT_FAMILY, 13)              # 按钮文本
SMALL_FONT = (FONT_FAMILY, 11)               # 辅助文本

# 窗口配置：更舒展的尺寸
WINDOW_SIZE_MAIN = "700x820"
WINDOW_SIZE_PROGRESS = "520x190"
WINDOW_SIZE_MAPPING = "860x700"


class ColumnMapper:
    """列映射管理器（功能不变）"""
    DEFAULT_MAPPING = {
        '发起人姓名': ['发起人姓名', '对接人'],
        '发起时间': ['发起时间', '创建时间'],
        '当前周': ['当前周'],
        '项目名称': ['项目名称'],
        '产品线': ['产品线', '产品'],
        '申请状态': ['申请状态', '当前进度'],
        '特制化比例': ['特制化比例(%)', '特制化比例'],
        '可常规化比例': ['可常规化比例(%)', '可常规化比例'],
        '建议报价元': ['建议报价(元)', '报价金额'],
        '定制内容': ['定制内容'],
        '软件版本': ['软件版本/产品名称', '产品名称'],
        '硬件情况': ['硬件情况（分辨率）/原产品主型号', '原产品主型号'],
        '销售部门': ['销售部门'],
        '定制人': ['定制人/销售经理', '销售经理']
    }
    
    OUTPUT_COLUMNS = {
        '发起人姓名': '对接人（发起人）',
        '发起时间': '发起时间',
        '当前周': '当前周',
        '项目名称': '项目名称',
        '产品线': '产品线',
        '申请状态': '当前进度',
        '特制化比例': '特制化比例(%)',
        '可常规化比例': '可常规化比例(%)',
        '建议报价元': '建议报价(元)',
        '定制内容': '定制内容',
        '软件版本': '软件版本/产品名称',
        '硬件情况': '硬件情况（分辨率）/原产品主型号',
        '销售部门': '销售部门',
        '定制人': '定制人/销售经理'
    }

    def __init__(self):
        self.load_mapping()
    
    def load_mapping(self):
        try:
            if os.path.exists('column_mapping.json'):
                with open('column_mapping.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.column_mapping = data.get('mapping', self.DEFAULT_MAPPING)
                    self.output_columns = data.get('output_columns', self.OUTPUT_COLUMNS)
            else:
                self.column_mapping = self.DEFAULT_MAPPING
                self.output_columns = self.OUTPUT_COLUMNS
        except Exception as e:
            print(f"加载配置失败: {e}")
            self.column_mapping = self.DEFAULT_MAPPING
            self.output_columns = self.OUTPUT_COLUMNS
    
    def save_mapping(self):
        try:
            with open('column_mapping.json', 'w', encoding='utf-8') as f:
                json.dump({
                    'mapping': self.column_mapping,
                    'output_columns': self.output_columns
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def get_mapping(self):
        return self.column_mapping

    def get_output_columns(self):
        return self.output_columns


# -------------------------- 核心数据处理函数（保持不变） --------------------------
def deep_clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_columns = []
    for col in df.columns:
        if str(col).startswith('Unnamed:'):
            if len(df) > 0:
                first_row_value = str(df.iloc[0][col]) if not pd.isna(df.iloc[0][col]) else ''
                if first_row_value and not first_row_value.startswith('Unnamed:'):
                    cleaned_columns.append(re.sub(r'[\s：()（）\n\t]', '', first_row_value).strip())
                else:
                    cleaned_columns.append(str(col))
            else:
                cleaned_columns.append(str(col))
        else:
            cleaned_columns.append(re.sub(r'[\s：()（）\n\t]', '', str(col)).strip())
    
    df.columns = cleaned_columns
    return df.dropna(how='all')


def dynamic_column_matching(df, column_mapper):
    column_mapping = column_mapper.get_mapping()
    matched = {}
    print("输入文件的列名：", df.columns.tolist())
    
    for target, aliases in column_mapping.items():
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
            print(f"警告：列[{target}]未找到\n尝试匹配的别名：{aliases}")
    
    return matched


def get_sheets_with_data(file_path):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheets_with_data = []
        
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=10)
                if not df.empty and len(df) > 0:
                    first_row = df.iloc[0].astype(str)
                    non_empty_count = first_row.count()
                    if non_empty_count >= 5:
                        header_keywords = ['时间', '日期', '申请', '审批', '金额', '报价', '产品', '类型']
                        first_row_text = ' '.join(first_row.tolist()).lower()
                        if any(keyword in first_row_text for keyword in header_keywords):
                            sheets_with_data.append(sheet_name)
                        elif len(df.columns) >= 10:
                            sheets_with_data.append(sheet_name)
            except Exception:
                continue
        
        return sheets_with_data
    except Exception as e:
        print(f"读取工作表列表失败: {e}")
        return []


def process_raw_excel(input_file, output_file, start_date=None, end_date=None, target_product=None, new_contact=None, progress_callback=None):
    try:
        if progress_callback:
            progress_callback(10, "正在分析文件结构...")
        
        sheet_names = get_sheets_with_data(input_file)
        if not sheet_names:
            raise Exception("未找到包含数据的工作表")
        
        if progress_callback:
            progress_callback(20, f"发现 {len(sheet_names)} 个工作表: {sheet_names}")
        
        all_data = []
        for i, sheet_name in enumerate(sheet_names):
            try:
                if progress_callback:
                    progress_callback(20 + i * 20 // len(sheet_names), f"正在读取工作表: {sheet_name}")
                
                df = pd.read_excel(
                    input_file, 
                    sheet_name=sheet_name,
                    header=1,
                    converters={'发起时间': str}
                )
                
                df['数据来源'] = sheet_name
                all_data.append(df)
            except Exception as e:
                print(f"读取工作表 {sheet_name} 失败: {e}")
                continue
        
        if not all_data:
            raise Exception("未能读取任何工作表数据")
        
        if progress_callback:
            progress_callback(40, "合并所有工作表数据...")
        
        combined_df = pd.concat(all_data, ignore_index=True)
        
        if progress_callback:
            progress_callback(50, f"数据合并完成，共 {len(combined_df)} 行记录")
        
        if progress_callback:
            progress_callback(60, "正在匹配列名...")
        column_mapper = ColumnMapper()
        matched = dynamic_column_matching(combined_df, column_mapper)
        
        if start_date and end_date:
            if progress_callback:
                progress_callback(70, f"筛选日期范围: {start_date} 至 {end_date}")
            
            try:
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df[time_column], 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                    
                    if combined_df['parsed_time'].isna().all():
                        date_pattern = r'(\d{4}-\d{2}-\d{2})'
                        combined_df['parsed_time'] = combined_df[time_column].apply(
                            lambda x: pd.to_datetime(re.search(date_pattern, str(x)).group(1)) if re.search(date_pattern, str(x)) else pd.NaT
                        )
                else:
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df.get('发起时间', pd.Series([pd.NaT] * len(combined_df))), 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                
                mask = (
                    (combined_df['parsed_time'] >= start_date) & 
                    (combined_df['parsed_time'] <= end_date)
                )
                filtered_df = combined_df[mask]
                
                if progress_callback:
                    progress_callback(80, f"日期筛选完成，剩余 {len(filtered_df)} 行记录")
            except Exception as e:
                print(f"日期筛选失败，将保留所有数据: {e}")
                filtered_df = combined_df
        else:
            if 'parsed_time' not in combined_df.columns:
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df[time_column], 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                else:
                    combined_df['parsed_time'] = pd.Series([pd.NaT] * len(combined_df))
            filtered_df = combined_df
        
        if progress_callback:
            progress_callback(90, "正在生成输出数据...")
        
        filtered_df.loc[:, '当前周'] = filtered_df['parsed_time'].dt.isocalendar().week
        
        desired_order = [
            '对接人（发起人）', '发起时间', '当前周', '项目名称', '产品线',
            '当前进度', '特制化比例(%)', '可常规化比例(%)', '建议报价(元)',
            '定制内容', '软件版本/产品名称', '硬件情况（分辨率）/原产品主型号',
            '销售部门', '定制人/销售经理'
        ]
        
        output_df = pd.DataFrame()
        
        column_mappings = {
            '对接人（发起人）': ['发起人姓名', '发起人姓名 ', '对接人'],
            '发起时间': ['发起时间'],
            '当前周': ['当前周'],
            '项目名称': ['项目名称'],
            '产品线': ['产品线'],
            '当前进度': ['申请状态', '当前进度'],
            '特制化比例(%)': ['特制化比例(%)', '特制化比例'],
            '可常规化比例(%)': ['可常规化比例(%)', '可常规化比例'],
            '建议报价(元)': ['建议报价(元)'],
            '定制内容': ['定制内容'],
            '软件版本/产品名称': ['软件版本/产品名称', '产品名称'],
            '硬件情况（分辨率）/原产品主型号': ['硬件情况（分辨率）/原产品主型号', '原产品主型号'],
            '销售部门': ['销售部门'],
            '定制人/销售经理': ['定制人/销售经理', '销售经理']
        }
        
        for target_col in desired_order:
            found = False
            for source_col in column_mappings.get(target_col, []):
                matched_source_col = None
                for col in filtered_df.columns:
                    col_clean = re.sub(r'[\s：()（）\n\t]', '', str(col)).strip()
                    source_clean = re.sub(r'[\s：()（）\n\t]', '', str(source_col)).strip()
                    if col_clean == source_clean:
                        matched_source_col = col
                        break
                
                if matched_source_col and matched_source_col in filtered_df.columns:
                    output_df[target_col] = filtered_df[matched_source_col]
                    found = True
                    break
            
            if not found:
                output_df[target_col] = ""
                print(f"警告：列[{target_col}]未找到，将填充空字符串")
        
        for col in desired_order:
            if col not in output_df.columns:
                output_df[col] = ""
        
        output_df = output_df[desired_order]
        
        if target_product and new_contact:
            if '产品线' in output_df.columns and '对接人（发起人）' in output_df.columns:
                output_df.loc[output_df['产品线'] == target_product, '对接人（发起人）'] = new_contact
        
        if '发起时间' in output_df.columns:
            output_df['发起时间'] = pd.to_datetime(output_df['发起时间'], errors='coerce')
            output_df = output_df.sort_values(by='发起时间', ascending=False, na_position='last')
            print(f"已按发起时间降序排列，共 {len(output_df)} 条记录")
        elif 'parsed_time' in filtered_df.columns:
            output_df = output_df.iloc[filtered_df['parsed_time'].sort_values(ascending=False, na_position='last').index]
            print(f"已按解析时间降序排列，共 {len(output_df)} 条记录")
        else:
            print("警告：未找到时间列，数据将保持原始顺序")
        
        if progress_callback:
            progress_callback(95, f"正在保存结果到: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='处理结果')
            
            worksheet = writer.sheets['处理结果']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
                for cell in column:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        
        if progress_callback:
            progress_callback(100, "文件处理完成!")
        
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        traceback.print_exc()
        raise e


# -------------------------- 柔和简约风UI工具函数 --------------------------
def center_window(window, width=None, height=None):
    """窗口居中，无多余效果"""
    window.update_idletasks()
    if not width:
        width = window.winfo_width()
    if not height:
        height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def init_soft_style(style):
    """初始化柔和简约风格"""
    # 基础框架：无框无阴影，干净清爽
    style.configure('Soft.Frame.TFrame', 
                   background=BG_MAIN,
                   borderwidth=0)
    
    # 卡片容器：纯白背景+柔和阴影+统一圆角
    style.configure('Soft.Card.TLabelframe', 
                   background=BG_CARD,
                   borderwidth=0,
                   relief='flat',
                   padding=(PADDING_MD, PADDING_MD, PADDING_MD, PADDING_MD),
                   borderradius=BORDER_RADIUS)
    # 卡片标题：浅灰下划线，精致不厚重
    style.configure('Soft.Card.Label.TLabelframe.Label', 
                   background=BG_CARD,
                   font=SUBTITLE_FONT,
                   foreground=TEXT_PRIMARY,
                   padding=(0, 0, 0, PADDING_SM),
                   anchor='nw',
                   borderwidth=0,
                   borderbottomwidth=1,
                   bordercolor=BORDER_BASE)
    
    # 按钮：轻盈风格，无厚重边框
    style.configure('Soft.Button.TButton', 
                   padding=(PADDING_MD, PADDING_SM), 
                   relief='flat', 
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=BORDER_RADIUS)
    
    # 主按钮：柔和主色，hover变浅
    style.configure('Soft.Button.Primary.TButton', 
                   background=PRIMARY_COLOR,
                   foreground='white')
    style.map('Soft.Button.Primary.TButton', 
              background=[('active', PRIMARY_LIGHT), ('pressed', PRIMARY_DARK)],
              relief=[('pressed', 'flat')])
    
    # 成功按钮：柔和绿色
    style.configure('Soft.Button.Success.TButton', 
                   background=SUCCESS_COLOR,
                   foreground='white')
    style.map('Soft.Button.Success.TButton', 
              background=[('active', SUCCESS_LIGHT), ('pressed', SUCCESS_DARK)],
              relief=[('pressed', 'flat')])
    
    # 危险按钮：柔和红色
    style.configure('Soft.Button.Danger.TButton', 
                   background=DANGER_COLOR,
                   foreground='white')
    style.map('Soft.Button.Danger.TButton', 
              background=[('active', DANGER_LIGHT), ('pressed', DANGER_DARK)],
              relief=[('pressed', 'flat')])
    
    # 文本按钮：无背景，仅文字颜色
    style.configure('Soft.Button.Text.TButton', 
                   background=BG_CARD,
                   foreground=PRIMARY_COLOR)
    style.map('Soft.Button.Text.TButton', 
              background=[('active', BG_INPUT)],
              relief=[('pressed', 'flat')])
    
    # 输入框：浅背景+细边框，聚焦变主色
    style.configure('Soft.Input.TEntry', 
                   padding=(PADDING_MD, PADDING_SM), 
                   relief='solid', 
                   borderwidth=1,
                   font=CONTENT_FONT,
                   foreground=TEXT_PRIMARY,
                   fieldbackground=BG_INPUT,
                   bordercolor=BORDER_BASE,
                   borderradius=BORDER_RADIUS)
    style.map('Soft.Input.TEntry', 
              bordercolor=[('focus', PRIMARY_COLOR), ('active', PRIMARY_COLOR)],
              relief=[('focus', 'solid')])
    
    # 标签：浅灰文本，不抢镜
    style.configure('Soft.Label.TLabel', 
                   background=BG_CARD, 
                   font=CONTENT_FONT,
                   foreground=TEXT_REGULAR,
                   padding=(0, 0, PADDING_LG, 0),
                   anchor='e')
    
    # 进度条：柔和圆角，低饱和主色
    style.configure('Soft.Progress.Horizontal.TProgressbar',
                   troughcolor=BG_INPUT,
                   background=PRIMARY_COLOR,
                   borderwidth=0,
                   borderradius=BORDER_RADIUS//2,
                   troughrelief='flat')
    
    # 标签页：简约下划线，无多余边框
    style.configure('Soft.Tabs.TNotebook',
                   background=BG_CARD,
                   borderwidth=0,
                   padding=(0, PADDING_SM))
    style.configure('Soft.Tabs.TNotebook.Tab',
                   background=BG_CARD,
                   font=CONTENT_FONT,
                   foreground=TEXT_SECONDARY,
                   padding=(PADDING_LG, PADDING_SM),
                   borderwidth=0,
                   borderradius=0)
    style.map('Soft.Tabs.TNotebook.Tab',
              background=[('selected', BG_CARD), ('active', BG_INPUT)],
              foreground=[('selected', PRIMARY_COLOR), ('active', PRIMARY_COLOR)],
              relief=[('selected', 'flat')])
    style.configure('Soft.Tabs.TNotebook.Frame',
                   borderwidth=0,
                   padding=(PADDING_MD, PADDING_MD, 0, 0))


# -------------------------- 柔和简约风主界面 --------------------------
def create_gui():
    """创建柔和简约风主界面"""
    root = tk.Tk()
    root.title("Excel数据处理工具")
    root.geometry(WINDOW_SIZE_MAIN)
    root.resizable(True, True)
    root.configure(bg=BG_MAIN)  # 柔和背景，不刺眼
    
    # 初始化柔和风格
    style = ttk.Style(root)
    init_soft_style(style)
    
    # 主框架：宽松内边距，呼吸感十足
    main_frame = ttk.Frame(root, padding=(PADDING_LG, PADDING_LG, PADDING_LG, PADDING_MD), style='Soft.Frame.TFrame')
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 标题区：简洁醒目，无多余装饰
    title_frame = ttk.Frame(main_frame, style='Soft.Frame.TFrame')
    title_frame.pack(fill=tk.X, pady=(0, PADDING_LG))
    title_label = ttk.Label(title_frame, 
                           text="Excel数据处理工具",
                           font=TITLE_FONT,
                           foreground=TEXT_PRIMARY,
                           background=BG_MAIN)
    title_label.pack(side=tk.LEFT)
    version_label = ttk.Label(title_frame,
                             text="v1.7",
                             font=SMALL_FONT,
                             foreground=TEXT_SECONDARY,
                             background=BG_MAIN)
    version_label.pack(side=tk.LEFT, padx=PADDING_SM, pady=(12, 0))
    
    # 变量定义
    input_entry = tk.StringVar()
    output_entry = tk.StringVar()
    start_date = tk.StringVar(value=datetime.now().replace(year=datetime.now().year-1).strftime("%Y/%m/%d"))
    end_date = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
    target_product_var = tk.StringVar()
    new_contact_var = tk.StringVar()
    
    # 日期快捷操作
    def set_week_start_end():
        today = datetime.now()
        week_start = today - timedelta(days=today.weekday())
        week_end = week_start + timedelta(days=6)
        start_date.set(week_start.strftime("%Y/%m/%d"))
        end_date.set(week_end.strftime("%Y/%m/%d"))
    
    def set_month_start_end():
        today = datetime.now()
        month_start = today.replace(day=1)
        next_month = month_start.replace(year=month_start.year + 1, month=1) if month_start.month == 12 else month_start.replace(month=month_start.month + 1)
        month_end = next_month - timedelta(days=1)
        start_date.set(month_start.strftime("%Y/%m/%d"))
        end_date.set(month_end.strftime("%Y/%m/%d"))
    
    def clear_dates():
        today = datetime.now()
        start_date.set(today.strftime("%Y/%m/%d"))
        end_date.set(today.strftime("%Y/%m/%d"))
    
    # 文件选择
    def select_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            input_entry.set(file_path)
            output_entry.set(os.path.dirname(file_path))
    
    def select_output_dir():
        dir_path = filedialog.askdirectory()
        if dir_path:
            output_entry.set(dir_path)
    
    # 进度窗口（柔和风格）
    def create_progress_window():
        progress_window = tk.Toplevel(root)
        progress_window.title("处理进度")
        progress_window.geometry(WINDOW_SIZE_PROGRESS)
        progress_window.resizable(False, False)
        progress_window.transient(root)
        progress_window.grab_set()
        progress_window.configure(bg=BG_CARD)
        
        # 进度窗口样式
        progress_style = ttk.Style(progress_window)
        init_soft_style(progress_style)
        
        # 进度窗口主框架：纯白背景+柔和阴影
        progress_frame = ttk.Frame(progress_window, padding=(PADDING_LG, PADDING_LG, PADDING_LG, PADDING_LG), style='Soft.Frame.TFrame')
        progress_frame.pack(fill=tk.BOTH, expand=True)
        progress_frame.configure(background=BG_CARD)
        
        # 进度提示：浅灰文本，不刺眼
        progress_label = ttk.Label(progress_frame, text="准备处理...", font=CONTENT_FONT, foreground=TEXT_REGULAR, background=BG_CARD)
        progress_label.pack(fill=tk.X, pady=(0, PADDING_MD))
        
        # 进度条：柔和圆角，主色填充
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_frame, 
                                      variable=progress_var, 
                                      maximum=100, 
                                      length=450, 
                                      style='Soft.Progress.Horizontal.TProgressbar')
        progress_bar.pack(fill=tk.X, pady=(0, PADDING_SM))
        
        # 进度更新
        def update_progress(progress, message):
            progress_var.set(progress)
            progress_label.config(text=message)
            progress_window.update()
        
        center_window(progress_window, width=520, height=190)
        return progress_window, update_progress
    
    # 开始处理
    def start_process():
        input_file = input_entry.get().strip()
        output_dir = output_entry.get().strip()
        
        if not input_file or not output_dir:
            messagebox.showerror("错误", "请选择输入文件和输出目录！")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "输入文件不存在！")
            return
        
        if not os.path.exists(output_dir):
            messagebox.showerror("错误", "输出目录不存在！")
            return
        
        # 禁用按钮防止重复点击
        process_btn.config(state='disabled')
        exit_btn.config(state='disabled')
        config_btn.config(state='disabled')
        
        # 创建进度窗口
        progress_window, update_progress = create_progress_window()
        
        try:
            start_dt = datetime.strptime(start_date.get(), "%Y/%m/%d")
            end_dt = datetime.strptime(end_date.get(), "%Y/%m/%d")
            target_product = target_product_var.get().strip() or None
            new_contact = new_contact_var.get().strip() or None
            
            output_file = f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
            
            success = process_raw_excel(
                input_file, output_file, start_dt, end_dt,
                target_product, new_contact,
                progress_callback=lambda p, msg: update_progress(p, msg)
            )
            
            if success:
                messagebox.showinfo("处理完成", f"文件处理成功！\n保存路径: {output_file}")
        except Exception as e:
            messagebox.showerror("处理失败", f"处理过程中发生错误: {str(e)}")
            traceback.print_exc()
        finally:
            progress_window.destroy()
            process_btn.config(state='normal')
            exit_btn.config(state='normal')
            config_btn.config(state='normal')
    
    # -------------------------- 柔和风格功能区块 --------------------------
    # 1. 文件设置（纯白卡片+柔和阴影）
    file_frame = ttk.LabelFrame(main_frame, 
                               text="文件设置", 
                               style='Soft.Card.TLabelframe')
    file_frame.pack(fill=tk.X, pady=(0, PADDING_MD))
    file_frame.configure(background=BG_CARD)
    # 卡片标题（带浅灰下划线）
    file_frame_label = ttk.Label(file_frame, text="文件设置", style='Soft.Card.Label.TLabelframe.Label')
    file_frame['labelwidget'] = file_frame_label
    file_frame_label.configure(background=BG_CARD)
    
    # 输入文件行（宽松间距）
    file_input_row = ttk.Frame(file_frame, style='Soft.Frame.TFrame')
    file_input_row.pack(fill=tk.X, pady=(0, PADDING_MD))
    file_input_row.configure(background=BG_CARD)
    file_input_label = ttk.Label(file_input_row, text="输入文件：", style='Soft.Label.TLabel')
    file_input_label.pack(side=tk.LEFT, anchor='center')
    file_input_label.configure(background=BG_CARD)
    file_input_entry = ttk.Entry(file_input_row, textvariable=input_entry, style='Soft.Input.TEntry')
    file_input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, PADDING_MD))
    file_input_btn = ttk.Button(file_input_row, text="浏览", command=select_input_file, style='Soft.Button.Primary.TButton')
    file_input_btn.pack(side=tk.LEFT, width=90)
    
    # 输出目录行
    file_output_row = ttk.Frame(file_frame, style='Soft.Frame.TFrame')
    file_output_row.pack(fill=tk.X)
    file_output_row.configure(background=BG_CARD)
    file_output_label = ttk.Label(file_output_row, text="输出目录：", style='Soft.Label.TLabel')
    file_output_label.pack(side=tk.LEFT, anchor='center')
    file_output_label.configure(background=BG_CARD)
    file_output_entry = ttk.Entry(file_output_row, textvariable=output_entry, style='Soft.Input.TEntry')
    file_output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, PADDING_MD))
    file_output_btn = ttk.Button(file_output_row, text="浏览", command=select_output_dir, style='Soft.Button.Primary.TButton')
    file_output_btn.pack(side=tk.LEFT, width=90)
    
    # 2. 日期筛选（纯白卡片）
    date_frame = ttk.LabelFrame(main_frame, 
                               text="日期筛选", 
                               style='Soft.Card.TLabelframe')
    date_frame.pack(fill=tk.X, pady=(0, PADDING_MD))
    date_frame.configure(background=BG_CARD)
    date_frame_label = ttk.Label(date_frame, text="日期筛选", style='Soft.Card.Label.TLabelframe.Label')
    date_frame['labelwidget'] = date_frame_label
    date_frame_label.configure(background=BG_CARD)
    
    # 日期选择行（宽松间距）
    date_select_row = ttk.Frame(date_frame, style='Soft.Frame.TFrame')
    date_select_row.pack(fill=tk.X, pady=(0, PADDING_MD))
    date_select_row.configure(background=BG_CARD)
    
    # 起始日期
    start_date_label = ttk.Label(date_select_row, text="起始日期：", style='Soft.Label.TLabel')
    start_date_label.pack(side=tk.LEFT, anchor='center')
    start_date_label.configure(background=BG_CARD)
    start_date_entry = DateEntry(date_select_row, 
                                textvariable=start_date, 
                                width=22, 
                                font=CONTENT_FONT,
                                foreground=TEXT_PRIMARY,
                                background=BG_INPUT,
                                date_pattern='yyyy/mm/dd',
                                locale='zh_CN',
                                borderwidth=1,
                                bordercolor=BORDER_BASE,
                                headersbackground=PRIMARY_COLOR,
                                selectbackground=PRIMARY_COLOR,
                                normalbackground=BG_CARD,
                                weekendbackground=BG_INPUT,
                                othermonthforeground=TEXT_SECONDARY,
                                othermonthbackground=BG_INPUT,
                                style='Soft.Input.TEntry')
    start_date_entry.pack(side=tk.LEFT, padx=(0, PADDING_LG))
    start_date_entry.configure(background=BG_INPUT)
    
    # 结束日期
    end_date_label = ttk.Label(date_select_row, text="结束日期：", style='Soft.Label.TLabel')
    end_date_label.pack(side=tk.LEFT, anchor='center')
    end_date_label.configure(background=BG_CARD)
    end_date_entry = DateEntry(date_select_row, 
                              textvariable=end_date, 
                              width=22, 
                              font=CONTENT_FONT,
                              foreground=TEXT_PRIMARY,
                              background=BG_INPUT,
                              date_pattern='yyyy/mm/dd',
                              locale='zh_CN',
                              borderwidth=1,
                              bordercolor=BORDER_BASE,
                              headersbackground=PRIMARY_COLOR,
                              selectbackground=PRIMARY_COLOR,
                              normalbackground=BG_CARD,
                              weekendbackground=BG_INPUT,
                              othermonthforeground=TEXT_SECONDARY,
                              othermonthbackground=BG_INPUT,
                              style='Soft.Input.TEntry')
    end_date_entry.pack(side=tk.LEFT)
    end_date_entry.configure(background=BG_INPUT)
    
    # 日期快捷按钮行（文本按钮，无背景）
    date_btn_row = ttk.Frame(date_frame, style='Soft.Frame.TFrame')
    date_btn_row.pack(fill=tk.X, anchor='w')
    date_btn_row.configure(background=BG_CARD)
    week_btn = ttk.Button(date_btn_row, text="本周", command=set_week_start_end, style='Soft.Button.Text.TButton', width=11)
    week_btn.pack(side=tk.LEFT, padx=(0, PADDING_MD))
    month_btn = ttk.Button(date_btn_row, text="本月", command=set_month_start_end, style='Soft.Button.Text.TButton', width=11)
    month_btn.pack(side=tk.LEFT, padx=(0, PADDING_MD))
    default_btn = ttk.Button(date_btn_row, text="恢复默认", command=clear_dates, style='Soft.Button.Text.TButton', width=11)
    default_btn.pack(side=tk.LEFT)
    
    # 3. 可选设置（纯白卡片）
    optional_frame = ttk.LabelFrame(main_frame, 
                                   text="可选设置", 
                                   style='Soft.Card.TLabelframe')
    optional_frame.pack(fill=tk.X, pady=(0, PADDING_MD))
    optional_frame.configure(background=BG_CARD)
    optional_frame_label = ttk.Label(optional_frame, text="可选设置", style='Soft.Card.Label.TLabelframe.Label')
    optional_frame['labelwidget'] = optional_frame_label
    optional_frame_label.configure(background=BG_CARD)
    
    # 产品线设置行
    product_row = ttk.Frame(optional_frame, style='Soft.Frame.TFrame')
    product_row.pack(fill=tk.X, pady=(0, PADDING_MD))
    product_row.configure(background=BG_CARD)
    product_label = ttk.Label(product_row, text="产品线名称：", style='Soft.Label.TLabel')
    product_label.pack(side=tk.LEFT, anchor='center')
    product_label.configure(background=BG_CARD)
    product_entry = ttk.Entry(product_row, textvariable=target_product_var, style='Soft.Input.TEntry')
    product_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # 对接人设置行
    contact_row = ttk.Frame(optional_frame, style='Soft.Frame.TFrame')
    contact_row.pack(fill=tk.X)
    contact_row.configure(background=BG_CARD)
    contact_label = ttk.Label(contact_row, text="新对接人：", style='Soft.Label.TLabel')
    contact_label.pack(side=tk.LEFT, anchor='center')
    contact_label.configure(background=BG_CARD)
    contact_entry = ttk.Entry(contact_row, textvariable=new_contact_var, style='Soft.Input.TEntry')
    contact_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # 4. 操作按钮区（宽松间距）
    action_frame = ttk.Frame(main_frame, style='Soft.Frame.TFrame')
    action_frame.pack(fill=tk.X, pady=(PADDING_LG, 0), anchor='e')
    action_frame.configure(background=BG_MAIN)
    
    config_btn = ttk.Button(action_frame, text="列映射配置", command=lambda: create_mapping_window(root), style='Soft.Button.Text.TButton', width=13)
    config_btn.pack(side=tk.LEFT, padx=(0, PADDING_LG))
    
    # 填充空间
    spacer = ttk.Frame(action_frame, style='Soft.Frame.TFrame')
    spacer.pack(side=tk.LEFT, fill=tk.X, expand=True)
    spacer.configure(background=BG_MAIN)
    
    exit_btn = ttk.Button(action_frame, text="退出", command=root.quit, style='Soft.Button.Danger.TButton', width=11)
    exit_btn.pack(side=tk.LEFT, padx=(0, PADDING_MD))
    
    process_btn = ttk.Button(action_frame, text="开始处理", command=start_process, style='Soft.Button.Success.TButton', width=13)
    process_btn.pack(side=tk.LEFT)
    
    # 窗口居中
    center_window(root, width=700, height=820)
    
    # 启动主循环
    root.mainloop()


# -------------------------- 柔和风格列映射配置窗口 --------------------------
def create_mapping_window(parent):
    """柔和简约风的列映射配置窗口"""
    mapping_window = tk.Toplevel(parent)
    mapping_window.title("列映射配置")
    mapping_window.geometry(WINDOW_SIZE_MAPPING)
    mapping_window.resizable(True, True)
    mapping_window.transient(parent)
    mapping_window.grab_set()
    mapping_window.configure(bg=BG_CARD)
    
    # 初始化样式
    style = ttk.Style(mapping_window)
    init_soft_style(style)
    
    # 加载列映射
    column_mapper = ColumnMapper()
    
    # 主框架（纯白背景）
    main_frame = ttk.Frame(mapping_window, padding=(PADDING_LG, PADDING_LG, PADDING_LG, PADDING_MD), style='Soft.Frame.TFrame')
    main_frame.pack(fill=tk.BOTH, expand=True)
    main_frame.configure(background=BG_CARD)
    
    # 窗口标题（简洁醒目）
    title_label = ttk.Label(main_frame, text="列映射配置", font=TITLE_FONT, foreground=TEXT_PRIMARY, background=BG_CARD)
    title_label.pack(fill=tk.X, pady=(0, PADDING_LG))
    
    # 标签页容器（简约风格）
    notebook = ttk.Notebook(main_frame, style='Soft.Tabs.TNotebook')
    notebook.pack(fill=tk.BOTH, expand=True, pady=(0, PADDING_MD))
    notebook.configure(background=BG_CARD)
    
    # -------------------------- 标签页1：列映射配置 --------------------------
    mapping_tab = ttk.Frame(notebook, style='Soft.Frame.TFrame')
    notebook.add(mapping_tab, text="列映射配置")
    mapping_tab.configure(background=BG_CARD)
    
    # 滚动容器（无框无阴影）
    mapping_scroll_frame = ttk.Frame(mapping_tab, style='Soft.Frame.TFrame')
    mapping_scroll_frame.pack(fill=tk.BOTH, expand=True)
    mapping_scroll_frame.configure(background=BG_CARD)
    
    mapping_canvas = tk.Canvas(mapping_scroll_frame, bg=BG_CARD, highlightthickness=0)
    mapping_scrollbar = ttk.Scrollbar(mapping_scroll_frame, orient="vertical", command=mapping_canvas.yview)
    mapping_content_frame = ttk.Frame(mapping_canvas, bg=BG_CARD, style='Soft.Frame.TFrame')
    
    mapping_content_frame.bind(
        "<Configure>",
        lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all"))
    )
    
    mapping_canvas.create_window((0, 0), window=mapping_content_frame, anchor="nw", width=800)
    mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
    
    mapping_canvas.pack(side="left", fill="both", expand=True)
    mapping_scrollbar.pack(side="right", fill="y")
    
    # 存储输入控件
    mapping_entries = {}
    
    # 创建映射配置项（宽松间距）
    mapping_data = column_mapper.get_mapping()
    for i, (target_col, aliases) in enumerate(mapping_data.items()):
        row_frame = ttk.Frame(mapping_content_frame, bg=BG_CARD, style='Soft.Frame.TFrame')
        row_frame.pack(fill=tk.X, pady=(0, PADDING_MD), padx=PADDING_XS)
        
        # 标签（右对齐，浅灰文本）
        label = ttk.Label(row_frame, text=f"{target_col}：", style='Soft.Label.TLabel', width=20, anchor='e')
        label.pack(side=tk.LEFT, padx=(0, PADDING_LG))
        label.configure(background=BG_CARD)
        
        # 输入框（浅背景，细边框）
        entry = ttk.Entry(row_frame, style='Soft.Input.TEntry')
        entry.insert(0, ", ".join(aliases))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        mapping_entries[target_col] = entry
    
    # -------------------------- 标签页2：输出列配置 --------------------------
    output_tab = ttk.Frame(notebook, style='Soft.Frame.TFrame')
    notebook.add(output_tab, text="输出列配置")
    output_tab.configure(background=BG_CARD)
    
    # 滚动容器
    output_scroll_frame = ttk.Frame(output_tab, style='Soft.Frame.TFrame')
    output_scroll_frame.pack(fill=tk.BOTH, expand=True)
    output_scroll_frame.configure(background=BG_CARD)
    
    output_canvas = tk.Canvas(output_scroll_frame, bg=BG_CARD, highlightthickness=0)
    output_scrollbar = ttk.Scrollbar(output_scroll_frame, orient="vertical", command=output_canvas.yview)
    output_content_frame = ttk.Frame(output_canvas, bg=BG_CARD, style='Soft.Frame.TFrame')
    
    output_content_frame.bind(
        "<Configure>",
        lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all"))
    )
    
    output_canvas.create_window((0, 0), window=output_content_frame, anchor="nw", width=800)
    output_canvas.configure(yscrollcommand=output_scrollbar.set)
    
    output_canvas.pack(side="left", fill="both", expand=True)
    output_scrollbar.pack(side="right", fill="y")
    
    # 存储输出列输入控件
    output_entries = {}
    
    # 创建输出列配置项
    output_data = column_mapper.get_output_columns()
    for i, (source_col, target_col) in enumerate(output_data.items()):
        row_frame = ttk.Frame(output_content_frame, bg=BG_CARD, style='Soft.Frame.TFrame')
        row_frame.pack(fill=tk.X, pady=(0, PADDING_MD), padx=PADDING_XS)
        
        # 标签
        label = ttk.Label(row_frame, text=f"{source_col}：", style='Soft.Label.TLabel', width=20, anchor='e')
        label.pack(side=tk.LEFT, padx=(0, PADDING_LG))
        label.configure(background=BG_CARD)
        
        # 输入框
        entry = ttk.Entry(row_frame, style='Soft.Input.TEntry')
        entry.insert(0, target_col)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        output_entries[source_col] = entry
    
    # -------------------------- 操作按钮区 --------------------------
    btn_frame = ttk.Frame(main_frame, style='Soft.Frame.TFrame')
    btn_frame.pack(fill=tk.X, anchor='e')
    btn_frame.configure(background=BG_CARD)
    
    # 保存配置
    def save_mapping_config():
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
        
        messagebox.showinfo("保存成功", "列映射配置已成功保存！")
        mapping_window.destroy()
    
    cancel_btn = ttk.Button(btn_frame, text="取消", command=mapping_window.destroy, style='Soft.Button.Text.TButton', width=11)
    cancel_btn.pack(side=tk.RIGHT, padx=(0, PADDING_MD))
    
    save_btn = ttk.Button(btn_frame, text="保存配置", command=save_mapping_config, style='Soft.Button.Primary.TButton', width=13)
    save_btn.pack(side=tk.RIGHT)
    
    # 窗口居中
    center_window(mapping_window, width=860, height=700)


if __name__ == "__main__":
    create_gui()