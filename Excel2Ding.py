#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel转钉钉格式转换工具

该工具用于将Excel文件转换为钉钉审批单格式，支持时间筛选和列名映射。

Features:
- 自动识别Excel中的列名并进行映射
- 支持时间范围筛选
- 自动计算当前周
- 支持多工作表合并
- GUI界面支持
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import re
from datetime import datetime, timedelta
import traceback

# 窗口尺寸配置
MAIN_WINDOW_SIZE = "500x400"
PROGRESS_WINDOW_SIZE = "400x150"
MAPPING_WINDOW_SIZE = "600x500"

# 默认配置
DEFAULT_MAPPING = {
    '发起人姓名': ['发起人姓名'],
    '发起时间': ['发起时间'],
    '项目名称': ['项目名称', '项目'],
    '产品线': ['产品线', '产品'],
    '建议报价元': ['建议报价(元)', '报价金额'],
    '申请状态': ['申请状态', '当前进度'],
    '特制化比例': ['特制化比例(%)', '特制化比例'],
    '可常规化比例': ['可常规化比例(%)', '可常规化比例'],
    '定制内容': ['定制内容'],
    '产品名称': ['产品名称'],
    '原产品主型号': ['原产品主型号'],
    '销售部门': ['销售部门'],
    '定制人': ['销售经理', '定制人']
}

OUTPUT_COLUMNS = {
    '发起人姓名': '对接人',
    '发起时间': '创建时间',
    '当前周': '当前周',
    '项目名称': '项目名称',
    '产品线': '产品',
    '申请状态': '当前进度',
    '建议报价元': '报价金额',
    '特制化比例': '特制化比例',
    '可常规化比例': '可常规化比例',
    '定制内容': '定制内容',
    '产品名称': '产品名称',
    '原产品主型号': '原产品主型号',
    '销售部门': '销售部门',
    '定制人': '定制人'
}


class ColumnMapper:
    """列名映射器"""
    
    DEFAULT_MAPPING = {
        '发起人姓名': ['发起人姓名'],
        '发起时间': ['发起时间'],
        '项目名称': ['项目名称', '项目'],
        '产品线': ['产品线', '产品'],
        '建议报价元': ['建议报价(元)', '报价金额'],
        '申请状态': ['申请状态', '当前进度'],
        '特制化比例': ['特制化比例(%)', '特制化比例'],
        '可常规化比例': ['可常规化比例(%)', '可常规化比例'],
        '定制内容': ['定制内容'],
        '产品名称': ['产品名称'],
        '原产品主型号': ['原产品主型号'],
        '销售部门': ['销售部门'],
        '定制人': ['销售经理', '定制人']
    }
    
    OUTPUT_COLUMNS = {
        '发起人姓名': '对接人',
        '发起时间': '创建时间',
        '当前周': '当前周',
        '项目名称': '项目名称',
        '产品线': '产品',
        '申请状态': '当前进度',
        '建议报价元': '报价金额',
        '特制化比例': '特制化比例',
        '可常规化比例': '可常规化比例',
        '定制内容': '定制内容',
        '产品名称': '产品名称',
        '原产品主型号': '原产品主型号',
        '销售部门': '销售部门',
        '定制人': '定制人'
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
                self.save_mapping()
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
    df.columns = [re.sub(r'[\s：()（）\n\t]', '', str(col)).strip() for col in df.columns]
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
        if not found:
            raise ValueError(f"列[{target}]未找到，当前列：{df.columns.tolist()}")
    
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


def process_raw_excel(input_file, output_file, start_date=None, end_date=None, progress_callback=None):
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
        
        # 按发起时间从大到小排序
        # 首先找到包含'发起时间'关键词的列
        time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
        if time_columns:
            time_column = time_columns[0]
            # 尝试解析时间列
            combined_df['parsed_time'] = pd.to_datetime(
                combined_df[time_column], 
                errors='coerce',
                infer_datetime_format=True
            )
            # 按时间排序
            combined_df = combined_df.sort_values('parsed_time', ascending=False)
        
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
        
        # 使用匹配的列创建输出DataFrame
        output_columns = list(matched.values())
        # 确保'当前周'列在输出列中
        if '当前周' not in output_columns:
            output_columns.append('当前周')
        output_df = filtered_df[output_columns].copy()
        
        # 应用列名映射
        output_df.rename(columns=matched, inplace=True)
        
        # 应用输出列名映射
        output_df.rename(columns=column_mapper.get_output_columns(), inplace=True)
        
        # 调整'当前周'列的位置到第3列（索引2）
        if '当前周' in output_df.columns:
            current_week_col = output_df.pop('当前周')
            output_df.insert(2, '当前周', current_week_col)
        
        # 保存结果
        if progress_callback:
            progress_callback(95, f"正在保存结果到: {output_file}")
        
        output_df.to_excel(output_file, index=False)
        
        if progress_callback:
            progress_callback(100, "文件处理完成!")
        
        return True
        
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        raise e


def process_excel(
    input_path: str,
    start_date: str,
    end_date: str,
    output_path: str,
    target_product: str = None,
    new_contact: str = None,
    progress_callback: callable = None
) -> bool:
    """处理Excel文件的主函数
    
    读取输入Excel文件，按照配置进行数据处理，并输出结果。
    
    Args:
        input_path: 输入Excel文件路径
        start_date: 开始日期，格式为'YYYY/MM/DD'
        end_date: 结束日期，格式为'YYYY/MM/DD'
        output_path: 输出Excel文件路径
        target_product: 可选，目标产品线名称
        new_contact: 可选，替换后的对接人
        progress_callback: 可选，进度回调函数，接收进度值(0-100)和状态消息
    
    Returns:
        bool: 处理成功返回True，失败返回False
    
    Raises:
        ValueError: 当列名匹配失败或日期格式错误时抛出
    """
    try:
        column_mapper = ColumnMapper()
        
        # 读取文件 (10%)
        if progress_callback:
            progress_callback(10, "正在读取文件...")
            
        # 使用 converters 参数来处理日期列
        converters = {'发起时间': lambda x: str(x)}  # 将发起时间列转换为字符串
        df = pd.read_excel(
            input_path, 
            engine='openpyxl',
            converters=converters,
            # 确保以文本格式读取日期列
            dtype={'发起时间': str}
        )
        df = df.dropna(how='all', axis=1)
        
        # 清洗列名 (20%)
        if progress_callback:
            progress_callback(20, "正在清洗数据...")
        df = deep_clean_columns(df)
        
        # 列匹配 (30%)
        if progress_callback:
            progress_callback(30, "正在匹配列名...")
        matched = dynamic_column_matching(df, column_mapper)
        
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
            raise ValueError("日期解析失败，请检查\"发起时间\"列是否为有效的 Excel 序列号格式")
        
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
        
        # 修改这里，使用 column_mapper 的输出列配置
        output_df = filtered[list(matched.values())].rename(columns=matched)
        output_df = output_df.rename(columns=column_mapper.output_columns)
        output_df.insert(2, '当前周', filtered['当前周'])
        output_df['创建时间'] = filtered['datetime_obj'].dt.strftime('%Y/%m/%d %H:%M')
        
        # 如果设置了产品线替换规则
        if target_product and new_contact:
            # 替换指定产品线对应的对接人
            output_df.loc[output_df['产品'] == target_product, '对接人'] = new_contact
        
        # 修改这里，使用 column_mapper 的输出列顺序
        final_columns = list(column_mapper.output_columns.values())
        output_df = output_df.reindex(columns=final_columns)
        
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


def create_progress_window(root: tk.Tk) -> tuple:
    """创建进度条弹窗
    
    创建一个模态进度条窗口，用于显示处理进度。
    
    Args:
        root: 主窗口实例
    
    Returns:
        tuple: 包含(progress_window, progress_var, progress_label)的元组
    """
    progress_window = tk.Toplevel(root)
    setup_window(progress_window, "处理进度", PROGRESS_WINDOW_SIZE)
    progress_window.transient(root)
    progress_window.grab_set()
    
     # 设置进度条窗口图标
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "Excel2Ding.ico")
        progress_window.iconbitmap(icon_path)
    except Exception as e:
        print(f"加载图标失败: {e}")

    # 居中显示
    progress_window.update_idletasks()
    width = progress_window.winfo_width()
    height = progress_window.winfo_height()
    x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
    y = (progress_window.winfo_screenheight() // 2) - (height // 2)
    progress_window.geometry(f"{width}x{height}+{x}+{y}")
    
    frame = ttk.Frame(progress_window, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        frame,
        variable=progress_var,
        maximum=100,
        mode='determinate',
        style='Modern.Horizontal.TProgressbar'
    )
    progress_bar.pack(fill=tk.X, pady=(0, 10))
    
    progress_label = ttk.Label(frame, 
                              text="⏳ 准备处理...",
                              style='TLabel',  # 添加这行
                              font=('Microsoft YaHei UI', 10))
    progress_label.pack(anchor="w")
    
    return progress_window, progress_var, progress_label


def center_window(window):
    """使窗口居中显示"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def set_window_icon(window):
    """设置窗口图标"""
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "Excel2Ding.ico")
        window.iconbitmap(icon_path)
    except Exception as e:
        print(f"加载图标失败: {e}")


def setup_window(window, title, size):
    """设置窗口基本属性"""
    window.title(title)
    window.geometry(size)
    window.resizable(False, False)
    set_window_icon(window)


def create_mapping_window(root, column_mapper):
    """创建列映射配置窗口"""
    mapping_window = tk.Toplevel(root)
    setup_window(mapping_window, "列映射配置", MAPPING_WINDOW_SIZE)
    mapping_window.transient(root)
    mapping_window.grab_set()
    
    # 创建主框架
    main_frame = ttk.Frame(mapping_window, padding=10)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 创建Notebook用于分隔映射配置和输出列配置
    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    
    # 映射配置标签页
    mapping_frame = ttk.Frame(notebook)
    notebook.add(mapping_frame, text="列映射配置")
    
    # 输出列配置标签页
    output_frame = ttk.Frame(notebook)
    notebook.add(output_frame, text="输出列配置")
    
    # 映射配置内容
    mapping_canvas = tk.Canvas(mapping_frame)
    mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=mapping_canvas.yview)
    mapping_scrollable_frame = ttk.Frame(mapping_canvas)
    
    mapping_scrollable_frame.bind(
        "<Configure>",
        lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all"))
    )
    
    mapping_canvas.create_window((0, 0), window=mapping_scrollable_frame, anchor="nw")
    mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
    
    mapping_canvas.pack(side="left", fill="both", expand=True)
    mapping_scrollbar.pack(side="right", fill="y")
    
    # 输出列配置内容
    output_canvas = tk.Canvas(output_frame)
    output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_canvas.yview)
    output_scrollable_frame = ttk.Frame(output_canvas)
    
    output_scrollable_frame.bind(
        "<Configure>",
        lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all"))
    )
    
    output_canvas.create_window((0, 0), window=output_scrollable_frame, anchor="nw")
    output_canvas.configure(yscrollcommand=output_scrollbar.set)
    
    output_canvas.pack(side="left", fill="both", expand=True)
    output_scrollbar.pack(side="right", fill="y")
    
    # 存储输入控件的字典
    mapping_entries = {}
    output_entries = {}
    
    # 创建映射配置输入框
    for i, (target, aliases) in enumerate(column_mapper.get_mapping().items()):
        ttk.Label(mapping_scrollable_frame, text=f"{target}:").grid(row=i, column=0, sticky="w", padx=(0, 10), pady=2)
        entry = ttk.Entry(mapping_scrollable_frame, width=50)
        entry.insert(0, ", ".join(aliases))
        entry.grid(row=i, column=1, sticky="ew", pady=2)
        mapping_entries[target] = entry
    
    # 配置列权重以使输入框可以扩展
    mapping_scrollable_frame.columnconfigure(1, weight=1)
    
    # 创建输出列配置输入框
    for i, (source, target) in enumerate(column_mapper.get_output_columns().items()):
        ttk.Label(output_scrollable_frame, text=f"{source}:").grid(row=i, column=0, sticky="w", padx=(0, 10), pady=2)
        entry = ttk.Entry(output_scrollable_frame, width=30)
        entry.insert(0, target)
        entry.grid(row=i, column=1, sticky="ew", pady=2)
        output_entries[source] = entry
    
    # 配置列权重以使输入框可以扩展
    output_scrollable_frame.columnconfigure(1, weight=1)
    
    # 保存按钮
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
    
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X)
    
    save_button = ttk.Button(button_frame, text="保存配置", command=save_config)
    save_button.pack(side=tk.RIGHT, padx=(10, 0))
    
    cancel_button = ttk.Button(button_frame, text="取消", command=mapping_window.destroy)
    cancel_button.pack(side=tk.RIGHT)


def create_gui():
    """创建GUI界面"""
    root = tk.Tk()
    setup_window(root, "Excel转钉钉格式转换工具", MAIN_WINDOW_SIZE)
    
    # 设置样式
    style = ttk.Style()
    style.theme_use('clam')
    
    # 创建主框架
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 变量定义
    input_file_var = tk.StringVar()
    output_file_var = tk.StringVar()
    start_date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
    end_date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
    target_product_var = tk.StringVar()
    new_contact_var = tk.StringVar()
    
    # 日期操作函数
    def add_days(var, days):
        try:
            date = datetime.strptime(var.get(), "%Y/%m/%d")
            new_date = date + timedelta(days=days)
            var.set(new_date.strftime("%Y/%m/%d"))
        except Exception as e:
            messagebox.showerror("错误", f"日期格式错误: {e}")
    
    def subtract_days(var, days):
        try:
            date = datetime.strptime(var.get(), "%Y/%m/%d")
            new_date = date - timedelta(days=days)
            var.set(new_date.strftime("%Y/%m/%d"))
        except Exception as e:
            messagebox.showerror("错误", f"日期格式错误: {e}")
    
    # 文件选择函数
    def select_input_file():
        file_path = filedialog.askopenfilename(
            title="选择输入Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file_path:
            input_file_var.set(file_path)
            # 自动生成输出文件名
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_dir = os.path.dirname(file_path)
            output_file = os.path.join(output_dir, f"{base_name}_处理结果.xlsx")
            output_file_var.set(output_file)
    
    def select_output_file():
        file_path = filedialog.asksaveasfilename(
            title="选择输出Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if file_path:
            output_file_var.set(file_path)
    
    # 处理开始函数
    def start_process():
        input_file = input_file_var.get()
        output_file = output_file_var.get()
        start_date = start_date_var.get()
        end_date = end_date_var.get()
        target_product = target_product_var.get().strip() or None
        new_contact = new_contact_var.get().strip() or None
        
        if not input_file:
            messagebox.showerror("错误", "请选择输入文件！")
            return
        
        if not output_file:
            messagebox.showerror("错误", "请选择输出文件！")
            return
        
        try:
            datetime.strptime(start_date, "%Y/%m/%d")
            datetime.strptime(end_date, "%Y/%m/%d")
        except Exception as e:
            messagebox.showerror("错误", f"日期格式错误: {e}")
            return
        
        # 创建进度条窗口
        progress_window, progress_var, progress_label = create_progress_window(root)
        
        def update_progress(value, message):
            progress_var.set(value)
            progress_label.config(text=message)
            progress_window.update_idletasks()
        
        # 在新线程中执行处理
        import threading
        def process_thread():
            try:
                success = process_raw_excel(input_file, output_file, start_date, end_date, update_progress)
                if success:
                    progress_window.destroy()
                    messagebox.showinfo("成功", f"文件处理完成！\n输出文件: {output_file}")
                else:
                    progress_window.destroy()
                    messagebox.showerror("错误", "文件处理失败！")
            except Exception as e:
                progress_window.destroy()
                messagebox.showerror("错误", f"处理过程中发生错误: {str(e)}")
                traceback.print_exc()
        
        threading.Thread(target=process_thread, daemon=True).start()
    
    # 使用提示
    ttk.Label(main_frame, text="使用提示：", font=('Microsoft YaHei UI', 10, 'bold')).pack(anchor="w", pady=(0, 5))
    hint_text = """
    1. 选择需要处理的Excel文件
    2. 设置筛选的日期范围
    3. 点击"开始处理"按钮
    4. 等待处理完成，结果将保存到指定位置
    """
    ttk.Label(main_frame, text=hint_text.strip(), justify=tk.LEFT).pack(anchor="w", pady=(0, 15))
    
    # 日期范围设置
    date_frame = ttk.LabelFrame(main_frame, text="日期范围设置", padding=10)
    date_frame.pack(fill=tk.X, pady=(0, 15))
    
    # 开始日期
    start_frame = ttk.Frame(date_frame)
    start_frame.pack(fill=tk.X, pady=(0, 5))
    ttk.Label(start_frame, text="开始日期:").pack(side=tk.LEFT)
    ttk.Entry(start_frame, textvariable=start_date_var).pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
    ttk.Button(start_frame, text="-7天", width=5, command=lambda: subtract_days(start_date_var, 7)).pack(side=tk.RIGHT, padx=(5, 0))
    ttk.Button(start_frame, text="-1天", width=5, command=lambda: subtract_days(start_date_var, 1)).pack(side=tk.RIGHT, padx=(5, 0))
    
    # 结束日期
    end_frame = ttk.Frame(date_frame)
    end_frame.pack(fill=tk.X)
    ttk.Label(end_frame, text="结束日期:").pack(side=tk.LEFT)
    ttk.Entry(end_frame, textvariable=end_date_var).pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
    ttk.Button(end_frame, text="+1天", width=5, command=lambda: add_days(end_date_var, 1)).pack(side=tk.RIGHT, padx=(5, 0))
    ttk.Button(end_frame, text="+7天", width=5, command=lambda: add_days(end_date_var, 7)).pack(side=tk.RIGHT, padx=(5, 0))
    
    # 文件设置
    file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding=10)
    file_frame.pack(fill=tk.X, pady=(0, 15))
    
    # 输入文件
    input_frame = ttk.Frame(file_frame)
    input_frame.pack(fill=tk.X, pady=(0, 5))
    ttk.Label(input_frame, text="输入文件:").pack(anchor="w")
    input_file_frame = ttk.Frame(input_frame)
    input_file_frame.pack(fill=tk.X, pady=(5, 0))
    ttk.Entry(input_file_frame, textvariable=input_file_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Button(input_file_frame, text="选择", command=select_input_file).pack(side=tk.RIGHT, padx=(5, 0))
    
    # 输出文件
    output_frame = ttk.Frame(file_frame)
    output_frame.pack(fill=tk.X)
    ttk.Label(output_frame, text="输出文件:").pack(anchor="w")
    output_file_frame = ttk.Frame(output_frame)
    output_file_frame.pack(fill=tk.X, pady=(5, 0))
    ttk.Entry(output_file_frame, textvariable=output_file_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Button(output_file_frame, text="选择", command=select_output_file).pack(side=tk.RIGHT, padx=(5, 0))
    
    # 产品线替换设置
    replace_frame = ttk.LabelFrame(main_frame, text="产品线对接人替换（可选）", padding=10)
    replace_frame.pack(fill=tk.X, pady=(0, 15))
    
    # 产品线名称
    product_frame = ttk.Frame(replace_frame)
    product_frame.pack(fill=tk.X, pady=(0, 5))
    ttk.Label(product_frame, text="产品线名称:").pack(side=tk.LEFT)
    ttk.Entry(product_frame, textvariable=target_product_var).pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
    
    # 新对接人
    contact_frame = ttk.Frame(replace_frame)
    contact_frame.pack(fill=tk.X)
    ttk.Label(contact_frame, text="新对接人:").pack(side=tk.LEFT)
    ttk.Entry(contact_frame, textvariable=new_contact_var).pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)
    
    # 操作按钮
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X)
    
    # 创建ColumnMapper实例
    column_mapper = ColumnMapper()
    
    # 配置按钮
    config_button = ttk.Button(button_frame, text="配置", command=lambda: create_mapping_window(root, column_mapper))
    config_button.pack(side=tk.LEFT)
    
    # 开始处理按钮
    start_button = ttk.Button(button_frame, text="开始处理", command=start_process)
    start_button.pack(side=tk.RIGHT)
    
    # 运行主循环
    root.mainloop()


if __name__ == "__main__":
    create_gui()
