#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据处理工具 GUI版本
基于V1.4核心逻辑，提供图形界面操作
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

# UI常量
WINDOW_SIZE = "900x600"
PADDING = 20

# 颜色方案
PRIMARY_COLOR = "#409EFF"  # Element UI 主色调
SUCCESS_COLOR = "#67C23A"  # 成功色
WARNING_COLOR = "#E6A23C"  # 警告色
DANGER_COLOR = "#F56C6C"   # 危险色
BG_COLOR = "#FFFFFF"      # 背景色改为白色，使界面更明亮
TEXT_COLOR = "#303133"     # 主文本色
BORDER_COLOR = "#E4E7ED"   # 边框色改为更柔和的灰色

# 字体配置 - 紧凑型设计
TITLE_FONT = ('Microsoft YaHei UI', 12, 'bold')
LABEL_FONT = ('Microsoft YaHei UI', 8)
BUTTON_FONT = ('Microsoft YaHei UI', 8)
ENTRY_FONT = ('Microsoft YaHei UI', 8)


class ColumnMapper:
    """列映射管理器"""
    
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
        """加载列映射配置"""
        try:
            if os.path.exists('column_mapping.json'):
                with open('column_mapping.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.column_mapping = data.get('mapping', self.DEFAULT_MAPPING)
                    self.output_columns = data.get('output_columns', self.OUTPUT_COLUMNS)
            else:
                self.column_mapping = self.DEFAULT_MAPPING
                self.output_columns = self.OUTPUT_COLUMNS
                # 不自动保存，避免覆盖自定义配置
                # self.save_mapping()
        except Exception as e:
            print(f"加载配置失败: {e}")
            self.column_mapping = self.DEFAULT_MAPPING
            self.output_columns = self.OUTPUT_COLUMNS
    
    def save_mapping(self):
        """保存列映射配置"""
        try:
            with open('column_mapping.json', 'w', encoding='utf-8') as f:
                json.dump({
                    'mapping': self.column_mapping,
                    'output_columns': self.output_columns
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def get_mapping(self):
        """获取当前映射配置"""
        return self.column_mapping

    def get_output_columns(self):
        """获取输出列配置"""
        return self.output_columns


def deep_clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """深度清洗DataFrame的列名
    
    移除列名中的空白字符和特殊字符，并删除全为空的列。
    
    Args:
        df: 需要处理的DataFrame对象
    
    Returns:
        DataFrame: 清洗后的DataFrame对象
    """
    # 处理列名，移除空白字符和特殊字符
    cleaned_columns = []
    for col in df.columns:
        # 如果是Unnamed列，尝试从第一行获取真实列名
        if str(col).startswith('Unnamed:'):
            # 尝试从第一行获取列名
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
    
    # 删除全为空的列
    return df.dropna(how='all')


def dynamic_column_matching(df, column_mapper):
    """精确列名匹配"""
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
        # 不再抛出异常，而是打印警告并继续处理
        if not found:
            print(f"警告：列[{target}]未找到\n尝试匹配的别名：{aliases}")
    
    return matched


def get_sheets_with_data(file_path):
    """获取包含数据的工作表列表"""
    try:
        # 读取所有工作表名
        excel_file = pd.ExcelFile(file_path)
        sheets_with_data = []
        
        # 检查每个工作表是否有数据
        for sheet_name in excel_file.sheet_names:
            try:
                # 读取前几行检查是否有数据
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=10)
                # 过滤掉明显不是数据表的工作表（如标题行过长或包含特定文本）
                if not df.empty and len(df) > 0:
                    # 检查第一行是否包含明显的标题特征
                    first_row = df.iloc[0].astype(str)
                    # 如果第一行有足够多的非空值，则认为是数据表
                    non_empty_count = first_row.count()
                    if non_empty_count >= 5:  # 至少有5列有数据才认为是数据表
                        # 检查是否包含常见的表头关键词
                        header_keywords = ['时间', '日期', '申请', '审批', '金额', '报价', '产品', '类型']
                        first_row_text = ' '.join(first_row.tolist()).lower()
                        if any(keyword in first_row_text for keyword in header_keywords):
                            sheets_with_data.append(sheet_name)
                        # 如果没有关键词但有足够多的列，也认为是数据表
                        elif len(df.columns) >= 10:  # 原始文件有很多列
                            sheets_with_data.append(sheet_name)
            except Exception:
                continue
        
        return sheets_with_data
    except Exception as e:
        print(f"读取工作表列表失败: {e}")
        return []


def process_raw_excel(input_file, output_file, start_date=None, end_date=None, target_product=None, new_contact=None, progress_callback=None):
    """处理原始Excel文件，自动处理多sheet和时间格式"""
    try:
        if progress_callback:
            progress_callback(10, "正在分析文件结构...")
        
        # 获取所有包含数据的工作表
        sheet_names = get_sheets_with_data(input_file)
        if not sheet_names:
            raise Exception("未找到包含数据的工作表")
        
        if progress_callback:
            progress_callback(20, f"发现 {len(sheet_names)} 个工作表: {sheet_names}")
        
        # 读取所有工作表数据
        all_data = []
        for i, sheet_name in enumerate(sheet_names):
            try:
                if progress_callback:
                    progress_callback(20 + i * 20 // len(sheet_names), f"正在读取工作表: {sheet_name}")
                
                # 读取数据，跳过第一行（筛选条件），使用第二行作为表头
                df = pd.read_excel(
                    input_file, 
                    sheet_name=sheet_name,
                    header=1,  # 使用第二行作为表头
                    converters={'发起时间': str}
                )
                
                # 添加工作表名列
                df['数据来源'] = sheet_name
                all_data.append(df)
            except Exception as e:
                print(f"读取工作表 {sheet_name} 失败: {e}")
                continue
        
        if not all_data:
            raise Exception("未能读取任何工作表数据")
        
        # 合并所有数据
        if progress_callback:
            progress_callback(40, "合并所有工作表数据...")
        
        combined_df = pd.concat(all_data, ignore_index=True)
        
        if progress_callback:
            progress_callback(50, f"数据合并完成，共 {len(combined_df)} 行记录")
        
        # 列匹配 (60%)
        if progress_callback:
            progress_callback(60, "正在匹配列名...")
        column_mapper = ColumnMapper()
        matched = dynamic_column_matching(combined_df, column_mapper)
        
        # 日期筛选逻辑
        if start_date and end_date:
            if progress_callback:
                progress_callback(70, f"筛选日期范围: {start_date} 至 {end_date}")
            
            try:
                # 查找包含'发起时间'关键词的列
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    
                    # 尝试从复杂的时间列中提取日期
                    # 首先尝试直接解析
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df[time_column], 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                    
                    # 如果直接解析失败，尝试从文本中提取日期
                    if combined_df['parsed_time'].isna().all():
                        # 使用正则表达式提取日期格式
                        date_pattern = r'(\d{4}-\d{2}-\d{2})'
                        combined_df['parsed_time'] = combined_df[time_column].apply(
                            lambda x: pd.to_datetime(re.search(date_pattern, str(x)).group(1)) if re.search(date_pattern, str(x)) else pd.NaT
                        )
                else:
                    # 如果没有找到包含'发起时间'的列，尝试使用'发起时间'列
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df.get('发起时间', pd.Series([pd.NaT] * len(combined_df))), 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                
                # 检查转换结果
                valid_dates = combined_df['parsed_time'].notna().sum()
                
                # 筛选日期范围
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
            # 即使没有日期筛选，也要确保有parsed_time列
            if 'parsed_time' not in combined_df.columns:
                # 尝试创建parsed_time列
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df[time_column], 
                        errors='coerce',
                        infer_datetime_format=True
                    )
                else:
                    # 如果没有找到时间列，创建全为NaT的列
                    combined_df['parsed_time'] = pd.Series([pd.NaT] * len(combined_df))
            filtered_df = combined_df
        
        # 生成输出数据 (90%)
        if progress_callback:
            progress_callback(90, "正在生成输出数据...")
        
        # 添加当前周列
        filtered_df.loc[:, '当前周'] = filtered_df['parsed_time'].dt.isocalendar().week
        
        # 定义期望的列顺序
        desired_order = [
            '对接人（发起人）',           # 对接人（发起人）
            '发起时间',                   # 发起时间
            '当前周',                     # 当前周
            '项目名称',                   # 项目名称
            '产品线',                     # 产品线
            '当前进度',                   # 当前进度
            '特制化比例(%)',              # 特制化比例
            '可常规化比例(%)',            # 可常规化比例
            '建议报价(元)',               # 建议报价(元)
            '定制内容',                   # 定制内容
            '软件版本/产品名称',          # 软件版本/产品名称
            '硬件情况（分辨率）/原产品主型号',  # 硬件情况（分辨率）/原产品主型号
            '销售部门',                   # 销售部门
            '定制人/销售经理'             # 定制人/销售经理
        ]
        
        # 创建输出DataFrame
        output_df = pd.DataFrame()
        
        # 直接映射所需的列名到输入文件中的对应列
        column_mappings = {
            '对接人（发起人）': ['发起人姓名', '发起人姓名 ', '对接人'],
            '发起时间': ['发起时间'],
            '当前周': ['当前周'],  # 这个是我们自己添加的
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
        
        # 为每个目标列寻找匹配的源列
        for target_col in desired_order:
            found = False
            for source_col in column_mappings.get(target_col, []):
                # 在原始数据框中查找匹配的列名（考虑可能的空格和特殊字符）
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
            
            # 如果没有找到匹配的列，填充空字符串
            if not found:
                output_df[target_col] = ""
                print(f"警告：列[{target_col}]未找到，将填充空字符串")
        
        # 确保所有期望的列都在输出DataFrame中
        for col in desired_order:
            if col not in output_df.columns:
                output_df[col] = ""  # 对于未找到的列，填充空字符串
        
        # 重新排列列的顺序
        output_df = output_df[desired_order]
        
        # 如果设置了产品线替换规则
        if target_product and new_contact:
            # 替换指定产品线对应的对接人
            if '产品线' in output_df.columns and '对接人（发起人）' in output_df.columns:
                output_df.loc[output_df['产品线'] == target_product, '对接人（发起人）'] = new_contact
        
        # 按【发起时间】或parsed_time降序排列数据
        if '发起时间' in output_df.columns:
            # 确保发起时间列是datetime类型
            output_df['发起时间'] = pd.to_datetime(output_df['发起时间'], errors='coerce')
            # 按发起时间降序排列，空值放在最后
            output_df = output_df.sort_values(by='发起时间', ascending=False, na_position='last')
            print(f"已按发起时间降序排列，共 {len(output_df)} 条记录")
        elif 'parsed_time' in filtered_df.columns:
            # 使用原始的parsed_time列进行排序
            output_df = output_df.iloc[filtered_df['parsed_time'].sort_values(ascending=False, na_position='last').index]
            print(f"已按解析时间降序排列，共 {len(output_df)} 条记录")
        else:
            print("警告：未找到时间列，数据将保持原始顺序")
        
        # 保存结果
        if progress_callback:
            progress_callback(95, f"正在保存结果到: {output_file}")
        
        # 保存到Excel文件
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='处理结果')
            
            # 调整列宽和对齐
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
                
                # 设置对齐
                for cell in column:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        
        if progress_callback:
            progress_callback(100, "文件处理完成!")
        
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        traceback.print_exc()
        raise e


def center_window(window):
    """使窗口居中显示"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def setup_window(window, title, size, resizable=(True, True)):
    """统一设置窗口属性"""
    window.title(title)
    window.geometry(size)
    window.configure(bg=BG_COLOR)
    window.resizable(*resizable)
    center_window(window)















def create_gui():
    """创建优化的GUI界面"""
    root = tk.Tk()
    root.title("Excel数据处理工具 v1.5")
    root.geometry("550x680")  # 紧凑型设计
    root.resizable(True, True)
    
    # 设置窗口图标
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "Excel2Ding.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"加载图标失败: {e}")
    
    # 基本样式设置
    style = ttk.Style()
    style.configure('TButton', 
                   padding=8, 
                   relief='flat', 
                   background=PRIMARY_COLOR,
                   foreground='black',
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=4)
    style.map('TButton', 
              background=[('active', PRIMARY_COLOR),
                         ('pressed', '#3A8EE6')],
              foreground=[('active', 'black'),
                         ('pressed', 'black')],
              relief=[('pressed', 'flat')])
    
    style.configure('TLabel', 
                   background=BG_COLOR, 
                   font=LABEL_FONT,
                   foreground=TEXT_COLOR,
                   padding=(2, 0))
    
    # 设置输入框样式
    style.configure('TEntry', 
                   padding=10, 
                   relief='solid', 
                   borderwidth=1,
                   font=ENTRY_FONT,
                   foreground=TEXT_COLOR)
    # 移除fieldbackground属性以避免与DateEntry组件冲突
    style.map('TEntry', 
              bordercolor=[('focus', PRIMARY_COLOR)],
              relief=[('focus', 'solid')])
    
    style.configure('TFrame', 
                   background=BG_COLOR,
                   borderwidth=0)
    
    style.configure('TLabelframe', 
                   background=BG_COLOR,
                   borderwidth=1,
                   relief='solid',
                   bordercolor=BORDER_COLOR,
                   padding=20)
    
    style.configure('TLabelframe.Label', 
                   background=BG_COLOR,
                   font=LABEL_FONT,
                   foreground=TEXT_COLOR,
                   padding=(5, 5, 5, 0))
    
    # 主框架 - 紧凑型设计
    main_frame = ttk.Frame(root, padding=15)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 标题 - 紧凑型设计
    title_label = ttk.Label(main_frame, 
                           text="Excel数据处理工具 v1.5",
                           font=TITLE_FONT,
                           foreground=PRIMARY_COLOR)
    title_label.pack(pady=(0, 15))
    
    # 定义变量
    input_entry = tk.StringVar()
    output_entry = tk.StringVar()
    start_date = tk.StringVar(value=datetime.now().replace(year=datetime.now().year-1).strftime("%Y/%m/%d"))
    end_date = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
    target_product_var = tk.StringVar()
    new_contact_var = tk.StringVar()
    
    # 日期操作函数
    def set_week_start_end():
        today = datetime.now()
        week_start = today - timedelta(days=today.weekday())
        week_end = week_start + timedelta(days=6)
        start_date.set(week_start.strftime("%Y/%m/%d"))
        end_date.set(week_end.strftime("%Y/%m/%d"))
    
    def set_month_start_end():
        today = datetime.now()
        month_start = today.replace(day=1)
        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)
        month_end = next_month - timedelta(days=1)
        start_date.set(month_start.strftime("%Y/%m/%d"))
        end_date.set(month_end.strftime("%Y/%m/%d"))
    
    def clear_dates():
        today = datetime.now()
        start_date.set(today.strftime("%Y/%m/%d"))
        end_date.set(today.strftime("%Y/%m/%d"))
    
    def select_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            input_entry.set(file_path)
            output_entry.set(os.path.dirname(file_path))
    
    def select_output_dir():
        dir_path = filedialog.askdirectory()
        if dir_path:
            output_entry.set(dir_path)
    
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
        
        process_btn.configure(state='disabled')
        exit_btn.configure(state='disabled')
        
        # 创建进度条窗口
        progress_window = tk.Toplevel(root)
        progress_window.title("处理进度")
        progress_window.geometry("400x120")
        progress_window.resizable(False, False)
        progress_window.transient(root)
        progress_window.grab_set()
        progress_window.configure(bg=BG_COLOR)
        
        # 进度窗口样式
        progress_style = ttk.Style()
        progress_style.configure('Custom.Horizontal.TProgressbar',
                                troughcolor='#E4E7ED',
                                background=PRIMARY_COLOR,
                                borderwidth=0,
                                borderradius=10)
        
        progress_frame = ttk.Frame(progress_window, padding=25)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        # 进度条
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=450, style='Custom.Horizontal.TProgressbar')
        progress_bar.pack(pady=(0, 15))
        
        # 进度标签
        progress_label = ttk.Label(progress_frame, text="准备处理...", font=LABEL_FONT, foreground=TEXT_COLOR)
        progress_label.pack()
        
        def update_progress(progress, message):
            progress_var.set(progress)
            progress_label.config(text=message)
            progress_window.update()
        
        try:
            # 解析日期
            start_dt = datetime.strptime(start_date.get(), "%Y/%m/%d")
            end_dt = datetime.strptime(end_date.get(), "%Y/%m/%d")
            target_product = target_product_var.get().strip() or None
            new_contact = new_contact_var.get().strip() or None
            
            # 生成输出文件路径
            output_file = f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
            
            # 执行处理
            success = process_raw_excel(
                input_file, output_file, start_dt, end_dt,
                target_product, new_contact,
                progress_callback=lambda p, msg: update_progress(p, msg)
            )
            
            if success:
                # 创建自定义对话框
                result_window = tk.Toplevel(root)
                result_window.title("处理完成")
                result_window.geometry("400x150")
                result_window.resizable(False, False)
                
                # 居中显示
                result_window.update_idletasks()
                width = result_window.winfo_width()
                height = result_window.winfo_height()
                x = (result_window.winfo_screenwidth() // 2) - (width // 2)
                y = (result_window.winfo_screenheight() // 2) - (height // 2)
                result_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
                
                # 设置窗口样式
                result_frame = ttk.Frame(result_window, padding=20)
                result_frame.pack(fill=tk.BOTH, expand=True)
                
                # 添加消息标签
                ttk.Label(
                    result_frame, 
                    text=f"文件处理成功！\n保存路径: {os.path.basename(output_file)}",
                    font=LABEL_FONT, 
                    justify=tk.LEFT
                ).pack(fill=tk.X, pady=(0, 20))
                
                # 按钮区域
                button_frame = ttk.Frame(result_frame)
                button_frame.pack(fill=tk.X, side=tk.BOTTOM)
                
                # 为了右对齐按钮添加一个空白框架
                spacer = ttk.Frame(button_frame)
                spacer.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                # 确定按钮，增加左侧间距
                ttk.Button(
                    button_frame, 
                    text="确定", 
                    style='TButton',
                    command=result_window.destroy,
                    width=10
                ).pack(side=tk.RIGHT, padx=(0, 10))
                # 打开按钮
                ttk.Button(
                    button_frame, 
                    text="打开文件", 
                    style='TButton',
                    command=lambda: open_file(output_file),
                    width=12
                ).pack(side=tk.RIGHT, padx=(10, 10))   
                # 打开文件函数
                def open_file(file_path):
                    try:
                        os.startfile(file_path)
                    except Exception as e:
                        messagebox.showerror("错误", f"无法打开文件: {str(e)}")
                    finally:
                        result_window.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")
            traceback.print_exc()
        finally:
            # 恢复按钮状态
            progress_window.destroy()
            process_btn.configure(state='normal')
            exit_btn.configure(state='normal')
            config_btn.configure(state='normal')
    
    # 文件设置区域 - 现代化布局
    file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding=15)
    file_frame.pack(fill=tk.X, pady=(0, 20))

    # 为输入文件行添加一致的间距和对齐
    ttk.Label(file_frame, text="输入文件:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=0, column=0, sticky=tk.W, pady=(8, 10))
    input_entry_widget = ttk.Entry(file_frame, textvariable=input_entry, width=45, font=ENTRY_FONT)
    input_entry_widget.grid(row=0, column=1, sticky=tk.EW, padx=(8, 8), pady=(8, 10))
    browse_input_btn = ttk.Button(file_frame, text="浏览", style='Primary.TButton', command=select_input_file)
    browse_input_btn.grid(row=0, column=2, padx=(0, 5), pady=(8, 10))
    
    # 为输出目录行添加一致的间距和对齐
    ttk.Label(file_frame, text="输出目录:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=1, column=0, sticky=tk.W, pady=(5, 10))
    output_entry_widget = ttk.Entry(file_frame, textvariable=output_entry, width=45, font=ENTRY_FONT)
    output_entry_widget.grid(row=1, column=1, sticky=tk.EW, padx=(8, 8), pady=(5, 10))
    browse_output_btn = ttk.Button(file_frame, text="浏览", style='Primary.TButton', command=select_output_dir)
    browse_output_btn.grid(row=1, column=2, padx=(0, 5), pady=(5, 10))
    
    # 设置列权重以使输入框可以扩展
    file_frame.columnconfigure(1, weight=1)
    
    # 日期筛选区域 - 现代化布局
    date_frame = ttk.LabelFrame(main_frame, text="日期筛选", padding=15)
    date_frame.pack(fill=tk.X, pady=(0, 20))
    
    # 起始日期
    ttk.Label(date_frame, text="起始日期:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=0, column=0, sticky=tk.W, pady=(8, 10))
    start_date_entry = DateEntry(date_frame, textvariable=start_date, width=15, font=ENTRY_FONT, 
                                foreground=TEXT_COLOR, background="white", 
                                date_pattern='yyyy/mm/dd', locale='zh_CN',
                                borderwidth=2, headersbackground='#409EFF',
                                selectbackground='#409EFF', normalbackground='white',
                                weekendbackground='#F5F7FA', othermonthforeground='#A8ABB2',
                                othermonthbackground='#F5F7FA')
    start_date_entry.grid(row=0, column=1, sticky=tk.W, padx=(8, 15), pady=(8, 10))
    
    # 结束日期
    ttk.Label(date_frame, text="结束日期:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=0, column=2, sticky=tk.W, pady=(8, 10))
    end_date_entry = DateEntry(date_frame, textvariable=end_date, width=15, font=ENTRY_FONT, 
                              foreground=TEXT_COLOR, background="white", 
                              date_pattern='yyyy/mm/dd', locale='zh_CN',
                              borderwidth=2, headersbackground='#409EFF',
                              selectbackground='#409EFF', normalbackground='white',
                              weekendbackground='#F5F7FA', othermonthforeground='#A8ABB2',
                              othermonthbackground='#F5F7FA')
    end_date_entry.grid(row=0, column=3, sticky=tk.W, padx=(8, 15), pady=(8, 10))
    
    # 快捷按钮区域 - 使用网格布局
    button_frame = ttk.Frame(date_frame)
    button_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(5, 10))
    
    # 本周按钮
    week_btn = ttk.Button(button_frame, text="本周", style='Primary.TButton', command=set_week_start_end, width=10)
    week_btn.pack(side=tk.LEFT, padx=(5, 5))
    
    # 本月按钮
    month_btn = ttk.Button(button_frame, text="本月", style='Primary.TButton', command=set_month_start_end, width=10)
    month_btn.pack(side=tk.LEFT, padx=(5, 5))
    
    # 恢复默认按钮
    default_btn = ttk.Button(button_frame, text="恢复默认", style='Info.TButton', command=clear_dates, width=10)
    default_btn.pack(side=tk.LEFT, padx=(5, 5))
    
    # 可选设置区域 - 现代化布局
    optional_frame = ttk.LabelFrame(main_frame, text="可选设置", padding=15)
    optional_frame.pack(fill=tk.X, pady=(0, 20))
    
    # 产品线名称
    ttk.Label(optional_frame, text="产品线名称:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=0, column=0, sticky=tk.W, pady=(8, 10))
    product_entry_widget = ttk.Entry(optional_frame, textvariable=target_product_var, width=35, font=ENTRY_FONT)
    product_entry_widget.grid(row=0, column=1, sticky=tk.EW, padx=(8, 8), pady=(8, 10))
    
    # 新对接人
    ttk.Label(optional_frame, text="新对接人:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(row=1, column=0, sticky=tk.W, pady=(5, 10))
    contact_entry_widget = ttk.Entry(optional_frame, textvariable=new_contact_var, width=35, font=ENTRY_FONT)
    contact_entry_widget.grid(row=1, column=1, sticky=tk.EW, padx=(8, 8), pady=(5, 10))
    
    # 设置列权重以使输入框可以扩展
    optional_frame.columnconfigure(1, weight=1)
    
    # 操作按钮区域 - 现代化布局
    action_frame = ttk.Frame(main_frame)
    action_frame.pack(fill=tk.X, pady=(20, 5))
    
    # 创建ColumnMapper实例
    column_mapper = ColumnMapper()
    
    # 列映射配置按钮
    config_btn = ttk.Button(action_frame, text="列映射配置", 
                           style='Info.TButton',
                           command=lambda: create_mapping_window(root, column_mapper),
                           width=15)
    config_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    # 添加一个空白框架来辅助布局，确保左右两侧的按钮有合适的间距
    spacer_frame = ttk.Frame(action_frame)
    spacer_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # 退出按钮
    exit_btn = ttk.Button(action_frame, text="退出", 
                         style='Danger.TButton', 
                         command=root.quit,
                         width=10)
    exit_btn.pack(side=tk.RIGHT, padx=(10, 0))
    
    # 开始处理按钮
    process_btn = ttk.Button(action_frame, text="开始处理", 
                            style='Success.TButton', 
                            command=start_process,
                            width=15)
    process_btn.pack(side=tk.RIGHT, padx=(0, 10))
    
    # 启动主循环
    root.mainloop()


def create_mapping_window(root, column_mapper):
    """创建现代化的列映射配置窗口"""
    mapping_window = tk.Toplevel(root)
    mapping_window.title("列映射配置")
    mapping_window.geometry("700x550")  # 适当增大窗口尺寸，提供更好的空间布局
    mapping_window.resizable(True, True)
    mapping_window.transient(root)  # 设置为子窗口
    mapping_window.grab_set()  # 模态窗口
    mapping_window.configure(bg=BG_COLOR)
    
    # 设置样式
    style = ttk.Style()
    
    # 按钮样式
    style.configure('TButton', 
                   padding=6, 
                   relief='flat', 
                   background=PRIMARY_COLOR,
                   foreground='black',
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=6)
    style.map('TButton', 
              background=[('active', PRIMARY_COLOR),
                         ('pressed', '#3A8EE6')],
              foreground=[('active', 'black'),
                         ('pressed', 'black')],
              relief=[('pressed', 'flat')])
    
    # 创建主色调按钮样式
    style.configure('Primary.TButton', 
                   padding=8, 
                   relief='flat', 
                   background=PRIMARY_COLOR,
                   foreground='white',
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=6)
    style.map('Primary.TButton', 
              background=[('active', '#66B1FF'),
                         ('pressed', '#3A8EE6')],
              foreground=[('active', 'white'),
                         ('pressed', 'white')],
              relief=[('pressed', 'flat')])
    
    # 创建成功按钮样式
    style.configure('Success.TButton', 
                   padding=8, 
                   relief='flat', 
                   background=SUCCESS_COLOR,
                   foreground='white',
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=6)
    style.map('Success.TButton', 
              background=[('active', '#85CE61'),
                         ('pressed', '#529B2E')],
              foreground=[('active', 'white'),
                         ('pressed', 'white')],
              relief=[('pressed', 'flat')])
    
    # 创建危险按钮样式
    style.configure('Danger.TButton', 
                   padding=8, 
                   relief='flat', 
                   background=DANGER_COLOR,
                   foreground='white',
                   font=BUTTON_FONT,
                   borderwidth=0,
                   borderradius=6)
    style.map('Danger.TButton', 
              background=[('active', '#F78989'),
                         ('pressed', '#E64242')],
              foreground=[('active', 'white'),
                         ('pressed', 'white')],
              relief=[('pressed', 'flat')])
    
    # 配置输入框样式
    style.configure('TEntry', 
                   padding=10, 
                   relief='solid', 
                   borderwidth=1,
                   font=ENTRY_FONT,
                   foreground=TEXT_COLOR,
                   fieldbackground=BG_COLOR)
    style.map('TEntry', 
              bordercolor=[('focus', PRIMARY_COLOR)],
              relief=[('focus', 'solid')])
    
    # 配置标签样式
    style.configure('TLabel', 
                   background=BG_COLOR, 
                   font=LABEL_FONT,
                   foreground=TEXT_COLOR,
                   padding=(5, 3))
    
    # 配置框架样式
    style.configure('TFrame', 
                   background=BG_COLOR,
                   borderwidth=0)
    
    # 配置带标签的框架样式
    style.configure('TLabelframe', 
                   background=BG_COLOR,
                   borderwidth=1,
                   relief='solid',
                   bordercolor=BORDER_COLOR,
                   padding=15)
    
    style.configure('TLabelframe.Label', 
                   background=BG_COLOR,
                   font=SUBTITLE_FONT,
                   foreground=TEXT_COLOR,
                   padding=(5, 5, 5, 0))
    
    # 主框架 - 现代化布局
    main_frame = ttk.Frame(mapping_window, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 标题
    title_label = ttk.Label(main_frame, 
                           text="列映射配置",
                           font=SUBTITLE_FONT,
                           foreground=PRIMARY_COLOR)
    title_label.pack(pady=(0, 15))
    
    # 创建Notebook用于分隔映射配置和输出列配置
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
    
    # 映射配置标签页
    mapping_frame = ttk.Frame(notebook)
    notebook.add(mapping_frame, text="列映射配置")
    
    # 输出列配置标签页
    output_frame = ttk.Frame(notebook)
    notebook.add(output_frame, text="输出列配置")
    
    # 映射配置内容 - 带滚动条
    mapping_canvas = tk.Canvas(mapping_frame, highlightthickness=0, bg=BG_COLOR)
    mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=mapping_canvas.yview)
    mapping_scrollable_frame = ttk.Frame(mapping_canvas, style='TFrame')
    
    mapping_scrollable_frame.bind(
        "<Configure>",
        lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all"))
    )
    
    mapping_canvas.create_window((0, 0), window=mapping_scrollable_frame, anchor="nw", width=640)  # 设置宽度以适应窗口
    mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
    
    mapping_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    mapping_scrollbar.pack(side="right", fill="y")
    
    # 输出列配置内容 - 带滚动条
    output_canvas = tk.Canvas(output_frame, highlightthickness=0, bg=BG_COLOR)
    output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_canvas.yview)
    output_scrollable_frame = ttk.Frame(output_canvas, style='TFrame')
    
    output_scrollable_frame.bind(
        "<Configure>",
        lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all"))
    )
    
    output_canvas.create_window((0, 0), window=output_scrollable_frame, anchor="nw", width=640)  # 设置宽度以适应窗口
    output_canvas.configure(yscrollcommand=output_scrollbar.set)
    
    output_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    output_scrollbar.pack(side="right", fill="y")
    
    # 存储输入控件的字典
    mapping_entries = {}
    output_entries = {}
    
    # 创建映射配置输入框 - 现代化布局
    for i, (target, aliases) in enumerate(column_mapper.get_mapping().items()):
        # 目标列标签
        ttk.Label(mapping_scrollable_frame, text=f"{target}:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(
            row=i, column=0, sticky="w", padx=(15, 15), pady=(10, 10))
        
        # 输入框
        entry = ttk.Entry(mapping_scrollable_frame, width=45, font=ENTRY_FONT)
        entry.insert(0, ", ".join(aliases))
        entry.grid(row=i, column=1, sticky="ew", pady=(10, 10), padx=(0, 15))
        mapping_entries[target] = entry
    
    # 配置列权重以使输入框可以扩展
    mapping_scrollable_frame.columnconfigure(1, weight=1)
    
    # 创建输出列配置输入框 - 现代化布局
    for i, (source, target) in enumerate(column_mapper.get_output_columns().items()):
        # 源列标签
        ttk.Label(output_scrollable_frame, text=f"{source}:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(
            row=i, column=0, sticky="w", padx=(15, 15), pady=(10, 10))
        
        # 输入框
        entry = ttk.Entry(output_scrollable_frame, width=30, font=ENTRY_FONT)
        entry.insert(0, target)
        entry.grid(row=i, column=1, sticky="ew", pady=(10, 10), padx=(0, 15))
        output_entries[source] = entry
    
    # 配置列权重以使输入框可以扩展
    output_scrollable_frame.columnconfigure(1, weight=1)
    
    # 保存按钮回调函数
    def save_config():
        # 更新映射配置
        new_mapping = {}
        for target, entry in mapping_entries.items():
            aliases = [alias.strip() for alias in entry.get().split(",") if alias.strip()]
            new_mapping[target] = aliases if aliases else [target]
        
        # 更新输出列配置
        new_output_columns = {}
        for source, entry in output_entries.items():
            new_output_columns[source] = entry.get().strip() or source
        
        # 保存到column_mapper
        column_mapper.column_mapping = new_mapping
        column_mapper.output_columns = new_output_columns
        column_mapper.save_mapping()
        
        messagebox.showinfo("成功", "配置已保存！")
        mapping_window.destroy()
    
    # 按钮区域 - 现代化布局
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(20, 5))
    
    # 添加一个空白框架来辅助布局
    spacer_frame = ttk.Frame(button_frame)
    spacer_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # 取消按钮
    ttk.Button(button_frame, text="取消", 
               style='Danger.TButton', 
               command=mapping_window.destroy,
               width=10).pack(side=tk.RIGHT, padx=(0, 10))
    
    # 保存配置按钮
    ttk.Button(button_frame, text="保存配置", 
               style='Success.TButton', 
               command=save_config,
               width=12).pack(side=tk.RIGHT)


if __name__ == "__main__":
    create_gui()