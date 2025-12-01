#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据处理工具 GUI版本
基于V1.5核心逻辑，提供图形界面操作
增加多产品线名称修改（原只有一个产品线名称的对应发起人的修改）
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
try:
    import ttkbootstrap as tb
except ImportError:
    import sys, subprocess
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "ttkbootstrap"], check=True)
        import ttkbootstrap as tb
    except Exception:
        tb = None

if tb:
    from ttkbootstrap.widgets import DateEntry as TBDateEntry
    _bootstrap_available = True
else:
    _bootstrap_available = False
    TBDateEntry = None
from tkcalendar import DateEntry
import json
from openpyxl.styles import Alignment
from openpyxl import load_workbook
import logging
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "V1.7"))
from ui_config import (
    WINDOW_SIZE,
    PADDING,
    PRIMARY_COLOR,
    SECONDARY_COLOR,
    SUCCESS_COLOR,
    WARNING_COLOR,
    DANGER_COLOR,
    BG_COLOR,
    TEXT_COLOR,
    SECONDARY_TEXT,
    BORDER_COLOR,
    PANEL_BG,
    CARD_BG,
    HOVER_BG,
    PLACEHOLDER_COLOR,
    TITLE_FONT,
    SUBTITLE_FONT,
    LABEL_FONT,
    BUTTON_FONT,
    ENTRY_FONT,
    INFO_COLOR,
    LIGHT_BORDER,
    FOCUS_BORDER,
    BUTTON_PADDING,
    ENTRY_PADDING,
    CARD_PADDING,
    GROUP_PADDING,
    apply_design_system,
    SHADOW_COLOR,
)
from core.state import AppState
from ui.widgets import make_button, make_date_entry
from core import transform as transform_core
from core import mapping as mapping_core
from ui import components as ui_components


warnings.filterwarnings('ignore')
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s'
)


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


def process_raw_excel(input_file, output_file, start_date=None, end_date=None, target_product=None, new_contact=None, product_contact_list=None, progress_callback=None, cancel_event=None):
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
            if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
                return False
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
                df = transform_core.deep_clean_columns(df)
                
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
        
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
        combined_df = pd.concat(all_data, ignore_index=True)
        
        if progress_callback:
            progress_callback(50, f"数据合并完成，共 {len(combined_df)} 行记录")
        
        # 列匹配 (60%)
        if progress_callback:
            progress_callback(60, "正在匹配列名...")
        column_mapper = mapping_core.ColumnMapper()
        matched = transform_core.dynamic_column_matching(combined_df, column_mapper)
        
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
                        combined_df[time_column].astype(str),
                        errors='coerce'
                    )
                    
                    # 如果直接解析失败，尝试从文本中提取日期
                    if combined_df['parsed_time'].isna().all():
                        # 使用正则表达式提取日期格式
                        date_pattern = r'(\d{4}-\d{2}-\d{2})'
                        def _extract_ymd(text):
                            m = re.search(date_pattern, str(text))
                            return m.group(1) if m else None
                        combined_df['parsed_time'] = pd.to_datetime(
                            combined_df[time_column].map(_extract_ymd),
                            errors='coerce'
                        )
                else:
                    # 如果没有找到包含'发起时间'的列，尝试使用'发起时间'列
                    combined_df['parsed_time'] = pd.to_datetime(
                        combined_df.get('发起时间', pd.Series([pd.NaT] * len(combined_df))).astype(str),
                        errors='coerce'
                    )
                # 若仍无法解析，则遍历行内容提取首个日期
                if combined_df['parsed_time'].isna().all():
                    date_any_pattern = re.compile(r'(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?)')
                    vals = []
                    for _, row in combined_df.iterrows():
                        text_line = ' '.join([str(v) for v in row.values])
                        m = date_any_pattern.search(text_line)
                        vals.append(m.group(1) if m else None)
                    combined_df['parsed_time'] = pd.to_datetime(pd.Series(vals), errors='coerce')
                
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
                        combined_df[time_column].astype(str),
                        errors='coerce'
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
        
        # 创建输出DataFrame，融合动态匹配与别名回退，提高填充完整度
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
        output_df = pd.DataFrame()
        cm = column_mapper.get_output_columns()
        rev_cm = {v: k for k, v in cm.items()}  # 输出列名 -> 规范源名

        alias_mappings = {
            '对接人（发起人）': ['发起人姓名', '对接人'],
            '发起时间': ['发起时间', '创建时间'],
            '当前周': ['当前周'],
            '项目名称': ['项目名称', '项目'],
            '产品线': ['产品线', '产品'],
            '当前进度': ['申请状态', '当前进度'],
            '特制化比例(%)': ['特制化比例(%)', '特制化比例'],
            '可常规化比例(%)': ['可常规化比例(%)', '可常规化比例'],
            '建议报价(元)': ['建议报价(元)', '报价金额'],
            '定制内容': ['定制内容'],
            '软件版本/产品名称': ['软件版本/产品名称', '产品名称'],
            '硬件情况（分辨率）/原产品主型号': ['硬件情况（分辨率）/原产品主型号', '原产品主型号'],
            '销售部门': ['销售部门'],
            '定制人/销售经理': ['定制人/销售经理', '销售经理'],
        }

        def find_source_column(candidates):
            for source_col in candidates:
                source_clean = re.sub(r'[\s：()（）\n\t]', '', str(source_col)).strip()
                for col in filtered_df.columns:
                    col_clean = re.sub(r'[\s：()（）\n\t]', '', str(col)).strip()
                    if col_clean == source_clean:
                        return col
            return None

        # 逐列输出填充
        for out_col in desired_order:
            filled = False
            # 先尝试动态匹配（通过规范源名）
            if out_col in rev_cm:
                norm = rev_cm[out_col]
                if norm in matched and matched[norm] in filtered_df.columns:
                    output_df[out_col] = filtered_df[matched[norm]]
                    filled = True
            # 回退：别名匹配
            if not filled:
                src = find_source_column(alias_mappings.get(out_col, []))
                if src:
                    output_df[out_col] = filtered_df[src]
                    filled = True
            # 特例：当前周源于解析时间
            if not filled and out_col == '当前周':
                output_df[out_col] = filtered_df['parsed_time'].dt.isocalendar().week
                filled = True
            # 默认空字符串
            if not filled:
                output_df[out_col] = ""

        # 若发起时间缺失，则回填解析时间
        try:
            if ('发起时间' in output_df.columns) and (output_df['发起时间'].isna().all() or (output_df['发起时间'] == "").all()):
                output_df['发起时间'] = filtered_df['parsed_time']
        except Exception:
            pass
        
        # 如果提供了多组产品线和对接人列表
        if product_contact_list and isinstance(product_contact_list, list):
            # 若产品线列缺失或为空，且仅提供单一映射，则全局填充产品线
            if '产品线' in output_df.columns:
                prod_series = output_df['产品线'].astype(str).str.strip()
                all_empty = (prod_series == "").all()
                if all_empty and len(product_contact_list) == 1:
                    default_product, _default_contact = product_contact_list[0]
                    output_df['产品线'] = default_product
            # 替换每组指定产品线对应的对接人（忽略大小写/空白差异）
            if '产品线' in output_df.columns and '对接人（发起人）' in output_df.columns:
                for product, contact in product_contact_list:
                    mask = output_df['产品线'].astype(str).str.strip().str.lower() == str(product).strip().lower()
                    output_df.loc[mask, '对接人（发起人）'] = contact
        # 兼容旧的单一产品线替换方式
        elif target_product and new_contact:
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
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
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



class ExcelProcessor:
    """业务逻辑处理层"""

    @staticmethod
    def validate_mappings(mappings):
        """验证产品线-对接人映射列表"""
        allowed = re.compile(r"^[\u4e00-\u9fa5A-Za-z0-9 _\-/()]+$")
        seen = set()
        for product, contact in mappings:
            if not product or not contact:
                return False, "产品线和对接人均不能为空"
            if not allowed.match(product) or not allowed.match(contact):
                return False, "存在非法字符，请仅使用中英文、数字和常用符号"
            if product in seen:
                return False, f"重复的产品线: {product}"
            seen.add(product)
        return True, "OK"

    @staticmethod
    def process(input_file, output_file, start_dt, end_dt, product_contact_list, progress_callback, cancel_event=None):
        logging.info("开始处理: input=%s, output=%s, range=%s-%s, mappings=%s",
                     input_file, output_file, start_dt, end_dt, product_contact_list)
        return process_raw_excel(
            input_file,
            output_file,
            start_dt,
            end_dt,
            product_contact_list=product_contact_list,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
        )


class ProductLineManager:
    """多产品线输入组件"""

    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.rows = []
        self.frame.columnconfigure(1, weight=1)
        self.frame.columnconfigure(3, weight=1)

    def add_row(self, product="", contact=""):
        idx = len(self.rows)
        product_var = tk.StringVar(value=product)
        contact_var = tk.StringVar(value=contact)

        ttk.Label(self.frame, text="产品线名称:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(
            row=idx, column=0, sticky=tk.W, pady=(8, 10)
        )
        product_entry = ttk.Entry(self.frame, textvariable=product_var, width=30, font=ENTRY_FONT)
        product_entry.grid(row=idx, column=1, sticky=tk.EW, padx=(8, 8), pady=(8, 10))

        ttk.Label(self.frame, text="新对接人:", font=LABEL_FONT, foreground=TEXT_COLOR).grid(
            row=idx, column=2, sticky=tk.W, pady=(8, 10)
        )
        contact_entry = ttk.Entry(self.frame, textvariable=contact_var, width=30, font=ENTRY_FONT)
        contact_entry.grid(row=idx, column=3, sticky=tk.EW, padx=(8, 8), pady=(8, 10))

        delete_btn = ttk.Button(
            self.frame, text="删除", style='Danger.TButton', command=lambda i=idx: self.remove_row(i)
        )
        delete_btn.grid(row=idx, column=4, padx=(5, 0), pady=(8, 10))

        self.rows.append((product_var, contact_var, product_entry, contact_entry, delete_btn))

    def remove_row(self, idx):
        for widget in self.frame.grid_slaves(row=idx):
            widget.destroy()
        if 0 <= idx < len(self.rows):
            self.rows.pop(idx)
        # 重排索引
        for i, (p_var, c_var, p_entry, c_entry, btn) in enumerate(self.rows):
            for widget in self.frame.grid_slaves(row=i + 1):
                widget.grid(row=i)
            btn.config(command=lambda j=i: self.remove_row(j))

    def get_mappings(self):
        mappings = []
        for product_var, contact_var, *_ in self.rows:
            p = product_var.get().strip()
            c = contact_var.get().strip()
            if p and c:
                mappings.append((p, c))
        return mappings

    def load_from_file(self, path='product_mapping.json'):
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for item in data.get('mappings', []):
                    self.add_row(item.get('product', ''), item.get('contact', ''))
        except Exception as e:
            logging.warning("加载产品线映射失败: %s", e)

    def save_to_file(self, path='product_mapping.json'):
        try:
            data = {'mappings': [{'product': p, 'contact': c} for p, c in self.get_mappings()]}
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.warning("保存产品线映射失败: %s", e)


def create_gui():
    """创建优化的GUI界面"""
    if _bootstrap_available and tb is not None:
        root = tb.Window(themename='cosmo')
    else:
        root = tk.Tk()
    root.title("Excel数据处理工具 v1.6")
    root.geometry(WINDOW_SIZE)
    root.minsize(600, 560)
    try:
        root.maxsize(1200, root.winfo_screenheight())
    except Exception:
        pass
    root.resizable(True, True)
    
    # 设置窗口图标
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "Excel2Ding.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"加载图标失败: {e}")
    
    style = ttk.Style()
    apply_design_system(style)
    

    
    # 启动淡入动画，提升初始体验
    try:
        root.attributes('-alpha', 0.0)
        def _fade_in(step=0):
            if step <= 10:
                root.attributes('-alpha', step/10)
                root.after(20, lambda: _fade_in(step+1))
        _fade_in()
    except Exception:
        pass

    app_notebook = ttk.Notebook(root)
    app_notebook.pack(fill=tk.BOTH, expand=True)

    # 主内容页签
    main_tab = ttk.Frame(app_notebook)
    app_notebook.add(main_tab, text="数据处理")

    content_canvas = tk.Canvas(main_tab, highlightthickness=0, bg=BG_COLOR)
    v_scroll = tk.Scrollbar(main_tab, orient='vertical', command=content_canvas.yview, width=8)
    h_scroll = tk.Scrollbar(main_tab, orient='horizontal', command=content_canvas.xview, width=8)
    content_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
    content_canvas.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

    main_frame = ttk.Frame(content_canvas, padding=12)
    content_window = content_canvas.create_window((0, 0), window=main_frame, anchor='nw')
    def rounded_container(parent, radius=6, fill=CARD_BG, pad=16):
        c = tk.Canvas(parent, bg=BG_COLOR, highlightthickness=0)
        f = ttk.Frame(c, padding=pad)
        def draw(_=None):
            w = parent.winfo_width()
            if w <= 1:
                w = 700
            h = f.winfo_reqheight() + pad * 2
            c.configure(width=w, height=h)
            c.delete('all')
            r = radius
            x1 = 6
            y1 = 6
            x2 = w - 12
            y2 = h - 12
            c.create_rectangle(x1+3, y1+3, x2+3, y2+3, fill=SHADOW_COLOR, outline=SHADOW_COLOR)
            c.create_arc(x1, y1, x1 + 2*r, y1 + 2*r, start=90, extent=90, fill=fill, outline=fill)
            c.create_arc(x2 - 2*r, y1, x2, y1 + 2*r, start=0, extent=90, fill=fill, outline=fill)
            c.create_arc(x1, y2 - 2*r, x1 + 2*r, y2, start=180, extent=90, fill=fill, outline=fill)
            c.create_arc(x2 - 2*r, y2 - 2*r, x2, y2, start=270, extent=90, fill=fill, outline=fill)
            c.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=fill)
            c.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline=fill)
            c.create_window((pad, pad), window=f, anchor='nw')
        parent.bind('<Configure>', draw)
        c.pack(fill=tk.X, pady=(0, 16))
        return f

    # 鼠标滚轮支持（按需启用）
    mousewheel_bound = False
    def _on_mousewheel(event):
        delta = int(-1*(event.delta/120))
        content_canvas.yview_scroll(delta, 'units')
    def _set_mousewheel_binding(need_bind: bool):
        nonlocal mousewheel_bound
        if need_bind and not mousewheel_bound:
            content_canvas.bind_all('<MouseWheel>', _on_mousewheel)
            mousewheel_bound = True
        elif not need_bind and mousewheel_bound:
            content_canvas.unbind_all('<MouseWheel>')
            mousewheel_bound = False

    _update_job = None
    def _update_scroll_region():
        bbox = content_canvas.bbox('all')
        content_canvas.configure(scrollregion=bbox)
        if not bbox:
            v_scroll.pack_forget()
            h_scroll.pack_forget()
            return
        needs_v = (bbox[3] or 0) > content_canvas.winfo_height()
        needs_h = (bbox[2] or 0) > content_canvas.winfo_width()
        if needs_v:
            v_scroll.pack(fill=tk.Y, side=tk.RIGHT)
        else:
            v_scroll.pack_forget()
        if needs_h:
            h_scroll.pack(fill=tk.X, side=tk.BOTTOM)
        else:
            h_scroll.pack_forget()
        _set_mousewheel_binding(needs_v or needs_h)
        if not (needs_v or needs_h):
            try:
                content_canvas.yview_moveto(0)
                content_canvas.xview_moveto(0)
            except Exception:
                pass
        try:
            content_canvas.itemconfigure(content_window, width=content_canvas.winfo_width())
        except Exception:
            pass

    def _schedule_update(*_args):
        nonlocal _update_job
        if _update_job:
            content_canvas.after_cancel(_update_job)
        _update_job = content_canvas.after(60, _update_scroll_region)

    main_frame.bind('<Configure>', _schedule_update)
    content_canvas.bind('<Configure>', _schedule_update)
    def apply_responsive_styles(width):
        if width < 720:
            style.configure('TLabel', font=('Microsoft YaHei UI', 10))
            style.configure('TLabelframe.Label', font=('Microsoft YaHei UI', 12, 'bold'))
            style.configure('TButton', font=('Microsoft YaHei UI', 9, 'bold'))
            style.configure('TEntry', font=('Microsoft YaHei UI', 9))
        elif width < 1024:
            style.configure('TLabel', font=LABEL_FONT)
            style.configure('TLabelframe.Label', font=SUBTITLE_FONT)
            style.configure('TButton', font=BUTTON_FONT)
            style.configure('TEntry', font=ENTRY_FONT)
        else:
            style.configure('TLabel', font=('Microsoft YaHei UI', 12))
            style.configure('TLabelframe.Label', font=('Microsoft YaHei UI', 14, 'bold'))
            style.configure('TButton', font=('Microsoft YaHei UI', 10, 'bold'))
            style.configure('TEntry', font=('Microsoft YaHei UI', 10))
    def _schedule_responsive(*_args):
        w = root.winfo_width()
        apply_responsive_styles(w)
    root.bind('<Configure>', _schedule_responsive)
    # 初始时根据内容区域决定是否启用滚动
    _schedule_update()
    
    # 现代化标题设计 - 改进版
    title_frame = tk.Frame(main_frame, bg=BG_COLOR)
    title_frame.pack(fill=tk.X, pady=(0, 16))
    
    title_container = tk.Frame(title_frame, bg=BG_COLOR)
    title_container.pack(side=tk.LEFT)
    
    title_label = ttk.Label(title_container, text="Excel数据处理工具", font=('Microsoft YaHei UI', 24, 'bold'), foreground=PRIMARY_COLOR)
    title_label.pack(anchor=tk.W)
    
    subtitle_label = ttk.Label(title_container, text="智能化数据处理和报表生成", font=('Microsoft YaHei UI', 12), foreground=SECONDARY_TEXT)
    subtitle_label.pack(anchor=tk.W, pady=(8, 0))
    
    version_label = ttk.Label(title_frame, text="v1.6", font=('Microsoft YaHei UI', 12), foreground=PLACEHOLDER_COLOR)
    version_label.pack(side=tk.RIGHT, pady=(15, 0))
    try:
        root.bind('<Return>', lambda e: start_process())
        root.bind('<Escape>', lambda e: root.quit())
    except Exception:
        pass
    
    # 定义变量
    input_entry = tk.StringVar()
    output_entry = tk.StringVar()
    start_date_var = tk.StringVar(value=datetime.now().replace(year=datetime.now().year-1).strftime("%Y/%m/%d"))
    end_date_var = tk.StringVar(value=datetime.now().strftime("%Y/%m/%d"))
    app_state = AppState()
    processor = ExcelProcessor
    
    # 日期操作函数
    def set_week_start_end():
        today = datetime.now()
        week_start = today - timedelta(days=today.weekday())
        week_end = week_start + timedelta(days=6)
        set_start_date(week_start.strftime("%Y/%m/%d"))
        set_end_date(week_end.strftime("%Y/%m/%d"))
    
    def set_month_start_end():
        today = datetime.now()
        month_start = today.replace(day=1)
        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)
        month_end = next_month - timedelta(days=1)
        set_start_date(month_start.strftime("%Y/%m/%d"))
        set_end_date(month_end.strftime("%Y/%m/%d"))
    
    def clear_dates():
        today = datetime.now()
        set_start_date(today.strftime("%Y/%m/%d"))
        set_end_date(today.strftime("%Y/%m/%d"))
    
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
        # config_btn 已移除，无需禁用

        # 创建加载遮罩层（专业化加载状态指示器）
        overlay_win = tk.Toplevel(root)
        overlay_win.overrideredirect(True)
        overlay_win.attributes('-topmost', True)
        try:
            overlay_win.attributes('-alpha', 0.0)
        except Exception:
            pass
        # 覆盖根窗口区域
        def _place_overlay():
            x = root.winfo_rootx()
            y = root.winfo_rooty()
            w = root.winfo_width()
            h = root.winfo_height()
            overlay_win.geometry(f"{w}x{h}+{x}+{y}")
        _place_overlay()
        root.bind('<Configure>', lambda e: _place_overlay())
        mask = tk.Frame(overlay_win, bg='#000000')
        mask.pack(fill=tk.BOTH, expand=True)

        content_card = ttk.Frame(mask, padding=30)
        content_card.place(relx=0.5, rely=0.5, anchor='center')
        progress_var = tk.DoubleVar(value=0)
        progress_style = ttk.Style()
        progress_style.configure('Custom.Horizontal.TProgressbar', troughcolor='#E5E7EB', background=PRIMARY_COLOR, borderwidth=0)
        progress_bar = ttk.Progressbar(content_card, variable=progress_var, maximum=100, length=380, style='Custom.Horizontal.TProgressbar')
        progress_bar.pack(pady=(0, 12))
        progress_label = ttk.Label(content_card, text="准备处理...", font=LABEL_FONT, foreground=TEXT_COLOR)
        progress_label.pack()
        import threading
        cancel_event = threading.Event()
        cancel_btn = make_button(content_card, text="取消", command=lambda: cancel_event.set(), width=10, role='danger')
        cancel_btn.pack(pady=(8, 0))

        # 淡入遮罩
        try:
            def _overlay_fade(step=0):
                if step <= 10:
                    overlay_win.attributes('-alpha', step/12)
                    overlay_win.after(15, lambda: _overlay_fade(step+1))
            _overlay_fade()
        except Exception:
            pass

        def update_progress(progress, message):
            progress_var.set(progress)
            progress_label.config(text=message)
            overlay_win.update_idletasks()
        
        try:
            # 解析日期
            start_dt = datetime.strptime(get_start_date(), "%Y/%m/%d")
            end_dt = datetime.strptime(get_end_date(), "%Y/%m/%d")
            
            # 收集并验证产品线映射
            product_contact_list = pl_manager.get_mappings()
            valid, msg = processor.validate_mappings(product_contact_list)
            if not valid:
                messagebox.showerror("错误", msg)
                try:
                    # 轻量提示，提高可感知度
                    toast = tk.Toplevel(root)
                    toast.overrideredirect(True)
                    toast.attributes('-topmost', True)
                    toast.configure(bg='#111827')
                    lbl = tk.Label(toast, text=msg, fg='white', bg='#111827', font=('Microsoft YaHei UI', 10))
                    lbl.pack(padx=12, pady=8)
                    tw = lbl.winfo_reqwidth() + 24
                    th = lbl.winfo_reqheight() + 16
                    root.update_idletasks()
                    rx = root.winfo_rootx()
                    ry = root.winfo_rooty()
                    rw = root.winfo_width()
                    rh = root.winfo_height()
                    toast.geometry(f"{tw}x{th}+{rx + rw//2 - tw//2}+{ry + rh - th - 40}")
                    def _auto_close():
                        try:
                            toast.destroy()
                        except Exception:
                            pass
                    toast.after(1800, _auto_close)
                except Exception:
                    pass
                return
            
            # 生成输出文件路径
            output_file = f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx"
            
            # 执行处理
            success = processor.process(
                input_file,
                output_file,
                start_dt,
                end_dt,
                product_contact_list,
                progress_callback=lambda p, msg: update_progress(p, msg),
                cancel_event=cancel_event,
            )
            
            if success:
                try:
                    toast = tk.Toplevel(root)
                    toast.overrideredirect(True)
                    toast.attributes('-topmost', True)
                    toast.configure(bg='#16A34A')
                    lbl = tk.Label(toast, text=f"文件处理成功：{os.path.basename(output_file)}", fg='white', bg='#16A34A', font=('Microsoft YaHei UI', 10))
                    lbl.pack(padx=14, pady=10)
                    tw = lbl.winfo_reqwidth() + 28
                    th = lbl.winfo_reqheight() + 20
                    root.update_idletasks()
                    rx = root.winfo_rootx()
                    ry = root.winfo_rooty()
                    rw = root.winfo_width()
                    rh = root.winfo_height()
                    toast.geometry(f"{tw}x{th}+{rx + rw//2 - tw//2}+{ry + rh - th - 50}")
                    def _auto_close():
                        try:
                            toast.destroy()
                        except Exception:
                            pass
                    toast.after(2000, _auto_close)
                except Exception:
                    pass
                try:
                    os.startfile(output_file)
                except Exception:
                    pass
            else:
                try:
                    toast = tk.Toplevel(root)
                    toast.overrideredirect(True)
                    toast.attributes('-topmost', True)
                    toast.configure(bg='#9CA3AF')
                    lbl = tk.Label(toast, text="处理已取消", fg='white', bg='#9CA3AF', font=('Microsoft YaHei UI', 10))
                    lbl.pack(padx=12, pady=8)
                    tw = lbl.winfo_reqwidth() + 24
                    th = lbl.winfo_reqheight() + 16
                    root.update_idletasks()
                    rx = root.winfo_rootx()
                    ry = root.winfo_rooty()
                    rw = root.winfo_width()
                    rh = root.winfo_height()
                    toast.geometry(f"{tw}x{th}+{rx + rw//2 - tw//2}+{ry + rh - th - 40}")
                    def _auto_close():
                        try:
                            toast.destroy()
                        except Exception:
                            pass
                    toast.after(1800, _auto_close)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")
            traceback.print_exc()
        finally:
            # 关闭遮罩并恢复按钮状态
            try:
                overlay_win.destroy()
            except Exception:
                pass
            process_btn.configure(state='normal')
            exit_btn.configure(state='normal')
            # config_btn.configure(state='normal')  # 已移除
    
    # 文件设置区域 - 改进版现代化布局
    file_frame = ttk.LabelFrame(main_frame, text="▌文件设置", padding=16)
    file_frame.pack(fill=tk.X, pady=(0, 25))
    try:
        file_frame.configure(style='Card.TLabelframe')
    except Exception:
        pass

    # 输入文件行 - 更大的间距和更好的对齐
    ttk.Label(file_frame, text="输入文件:", style='Card.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(8, 12))
    input_entry_widget = ttk.Entry(file_frame, textvariable=input_entry, width=42)
    input_entry_widget.grid(row=0, column=1, sticky=tk.EW, padx=(12, 12), pady=(8, 12))
    browse_input_btn = ttk.Button(file_frame, text="浏览", command=select_input_file, style='Secondary.TButton')
    browse_input_btn.grid(row=0, column=2, pady=(8, 12))
    
    # 输出目录行 - 更大的间距和更好的对齐
    ttk.Label(file_frame, text="输出目录:", style='Card.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(0, 12))
    output_entry_widget = ttk.Entry(file_frame, textvariable=output_entry, width=42)
    output_entry_widget.grid(row=1, column=1, sticky=tk.EW, padx=(12, 12), pady=(0, 12))
    browse_output_btn = ttk.Button(file_frame, text="浏览", command=select_output_dir, style='Secondary.TButton')
    browse_output_btn.grid(row=1, column=2, pady=(0, 12))
    
    # 设置列权重以使输入框可以扩展
    file_frame.columnconfigure(1, weight=1)
    
    # 日期筛选区域 - 改进版现代化布局
    date_frame = ttk.LabelFrame(main_frame, text="▌日期筛选", padding=16)
    date_frame.pack(fill=tk.X, pady=(0, 25))
    try:
        date_frame.configure(style='Card.TLabelframe')
    except Exception:
        pass
    
    # 起始日期 - 更大的间距和更好的对齐
    ttk.Label(date_frame, text="起始日期:", style='Card.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(8, 12))
    start_date_entry = make_date_entry(date_frame, width=14, dateformat='%Y/%m/%d', bootstyle='success', firstweekday=6)
    start_date_entry.grid(row=0, column=1, sticky=tk.W, padx=(12, 20), pady=(8, 12))
    # 初始化起始日期显示
    try:
        start_date_entry.entry.delete(0, tk.END)
        start_date_entry.entry.insert(0, start_date_var.get())
    except Exception:
        pass
    def set_start_date(val):
        start_date_var.set(val)
        try:
            start_date_entry.entry.delete(0, tk.END)
            start_date_entry.entry.insert(0, val)
        except Exception:
            pass
    def get_start_date():
        try:
            return start_date_entry.entry.get()
        except Exception:
            return start_date_var.get()
    
    # 结束日期 - 更大的间距和更好的对齐
    ttk.Label(date_frame, text="结束日期:", style='Card.TLabel').grid(row=0, column=2, sticky=tk.W, pady=(8, 12))
    end_date_entry = make_date_entry(date_frame, width=14, dateformat='%Y/%m/%d', bootstyle='success', firstweekday=6)
    end_date_entry.grid(row=0, column=3, sticky=tk.W, padx=(12, 0), pady=(8, 12))
    try:
        end_date_entry.entry.delete(0, tk.END)
        end_date_entry.entry.insert(0, end_date_var.get())
    except Exception:
        pass
    def set_end_date(val):
        end_date_var.set(val)
        try:
            end_date_entry.entry.delete(0, tk.END)
            end_date_entry.entry.insert(0, val)
        except Exception:
            pass
    def get_end_date():
        try:
            return end_date_entry.entry.get()
        except Exception:
            return end_date_var.get()
    
    # 快捷按钮区域 - 更大的间距
    button_frame = ttk.Frame(date_frame)
    button_frame.grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(0, 8))
    
    # 快捷按钮 - 更大的间距
    week_btn = make_button(button_frame, text="本周", command=set_week_start_end, width=10, role='primary')
    week_btn.pack(side=tk.LEFT, padx=(0, 8))
    month_btn = make_button(button_frame, text="本月", command=set_month_start_end, width=10, role='primary')
    month_btn.pack(side=tk.LEFT, padx=(0, 8))
    default_btn = make_button(button_frame, text="恢复默认", command=clear_dates, width=10, role='primary')
    default_btn.pack(side=tk.LEFT)


    
    # 可选设置区域 - 改进版现代化布局
    optional_frame = ttk.LabelFrame(main_frame, text="▌可选设置", padding=16)
    optional_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 25))
    try:
        optional_frame.configure(style='Card.TLabelframe')
    except Exception:
        pass

    notebook = ttk.Notebook(optional_frame)
    notebook.pack(fill=tk.BOTH, expand=True)

    product_tab = ttk.Frame(notebook)
    notebook.add(product_tab, text="产品线映射")

    pl_manager = ui_components.ProductLineManager(product_tab)
    pl_manager.add_row()

    action_tab_frame = ttk.Frame(product_tab)
    action_tab_frame.pack(fill=tk.X)
    ttk.Button(action_tab_frame, text="添加产品线", command=pl_manager.add_row, style='Info.TButton').pack(pady=(15, 0), anchor='w')

    # 载入历史映射
    pl_manager.load_from_file()
    
    # 操作按钮区域 - 改进版现代化布局
    overlay = tk.Frame(root, bg=BG_COLOR, padx=16, pady=16)
    overlay.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor='se')
    
    # 创建ColumnMapper实例
    column_mapper = mapping_core.ColumnMapper()
    
    # 列映射配置按钮
    # 配置页签（在顶层 Notebook 中）
    mapping_tab = ttk.Frame(app_notebook)
    app_notebook.add(mapping_tab, text="列映射配置")
    build_mapping_content(mapping_tab, column_mapper)

    info_tab = ttk.Frame(app_notebook)
    app_notebook.add(info_tab, text="软件信息")
    info_frame = ttk.Frame(info_tab, padding=30)
    info_frame.pack(fill=tk.BOTH, expand=True)
    ttk.Label(info_frame, text="工具作者：July-Chen-JIE", style='Card.TLabel', font=('Microsoft YaHei UI', 12)).pack(anchor=tk.W, pady=(0, 10))
    ttk.Label(info_frame, text="版本：V1.6_20251127", style='Card.TLabel', font=('Microsoft YaHei UI', 12)).pack(anchor=tk.W, pady=(0, 10))
    ttk.Label(info_frame, text="更新日期：20251127", style='Card.TLabel', font=('Microsoft YaHei UI', 12)).pack(anchor=tk.W)

    # 在叠加层内放置右下角固定按钮 - 更大的间距和更好的对齐
    # config_btn = tb.Button(
    #     overlay,
    #     text="列映射配置",
    #     command=lambda: app_notebook.select(mapping_tab),
    #     width=15,
    #     bootstyle='info'
    # )
    exit_btn = make_button(overlay, text="退出", command=root.quit, width=14, role='danger')
    process_btn = make_button(overlay, text="开始处理", command=start_process, width=18, role='primary')
    # 排列按钮，保持25px间距
    # info_btn = tb.Button(overlay, text="软件信息", command=lambda: app_notebook.select(info_tab), width=12, bootstyle='info')
    process_btn.grid(row=0, column=3, padx=(25, 0), pady=(8, 8))
    exit_btn.grid(row=0, column=2, padx=(25, 0), pady=(8, 8))
    # config_btn.grid(row=0, column=1, padx=(25, 0))
    # info_btn.grid(row=0, column=0)
    overlay.lift()
    
    # 添加一个空白框架来辅助布局，确保左右两侧的按钮有合适的间距
    spacer_frame = ttk.Frame(action_tab_frame)
    spacer_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # 移除底部原操作按钮区域（改为固定叠加层），保留变量引用在 start_process 内使用
    
    # 启动主循环
    root.mainloop()


def build_mapping_content(container, column_mapper):
    main_frame = ttk.Frame(container, padding=35)
    main_frame.pack(fill=tk.BOTH, expand=True)

    title_label = ttk.Label(main_frame, text="列映射配置", font=('Microsoft YaHei UI', 18, 'bold'), foreground=PRIMARY_COLOR)
    title_label.pack(pady=(0, 30))

    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

    mapping_frame = ttk.Frame(notebook)
    notebook.add(mapping_frame, text="列映射配置")

    output_frame = ttk.Frame(notebook)
    notebook.add(output_frame, text="输出列配置")

    mapping_canvas = tk.Canvas(mapping_frame, highlightthickness=0, bg=BG_COLOR)
    mapping_scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=mapping_canvas.yview)
    mapping_scrollable_frame = ttk.Frame(mapping_canvas, style='TFrame')
    mapping_scrollable_frame.bind("<Configure>", lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all")))
    mapping_window_id = mapping_canvas.create_window((0, 0), window=mapping_scrollable_frame, anchor="nw")
    mapping_canvas.configure(yscrollcommand=mapping_scrollbar.set)
    mapping_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    mapping_scrollbar.pack(side="right", fill="y")
    mapping_frame.bind('<Configure>', lambda e: mapping_canvas.itemconfigure(mapping_window_id, width=mapping_frame.winfo_width()))

    output_canvas = tk.Canvas(output_frame, highlightthickness=0, bg=BG_COLOR)
    output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=output_canvas.yview)
    output_scrollable_frame = ttk.Frame(output_canvas, style='TFrame')
    output_scrollable_frame.bind("<Configure>", lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all")))
    output_window_id = output_canvas.create_window((0, 0), window=output_scrollable_frame, anchor="nw")
    output_canvas.configure(yscrollcommand=output_scrollbar.set)
    output_canvas.pack(side="left", fill="both", expand=True, padx=(0, 1))
    output_scrollbar.pack(side="right", fill="y")
    output_frame.bind('<Configure>', lambda e: output_canvas.itemconfigure(output_window_id, width=output_frame.winfo_width()))

    mapping_entries = {}
    output_entries = {}

    for i, (target, aliases) in enumerate(column_mapper.get_mapping().items()):
        ttk.Label(mapping_scrollable_frame, text=f"{target}:").grid(
            row=i, column=0, sticky="w", padx=(20, 20), pady=(15, 15)
        )
        entry = ttk.Entry(mapping_scrollable_frame, width=40)
        entry.insert(0, ", ".join(aliases))
        entry.grid(row=i, column=1, sticky="ew", pady=(15, 15), padx=(0, 20))
        mapping_entries[target] = entry
    mapping_scrollable_frame.columnconfigure(1, weight=1)

    for i, (source, target) in enumerate(column_mapper.get_output_columns().items()):
        ttk.Label(output_scrollable_frame, text=f"{source}:").grid(
            row=i, column=0, sticky="w", padx=(20, 20), pady=(15, 15)
        )
        entry = ttk.Entry(output_scrollable_frame, width=40)
        entry.insert(0, target)
        entry.grid(row=i, column=1, sticky="ew", pady=(15, 15), padx=(0, 20))
        output_entries[source] = entry
    output_scrollable_frame.columnconfigure(1, weight=1)

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

    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(30, 10))
    spacer_frame = ttk.Frame(button_frame)
    spacer_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    def import_config():
        from tkinter import filedialog
        try:
            path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
            if not path:
                return
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
        from tkinter import filedialog
        try:
            path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
            if not path:
                return
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

    # 页签切换动画：进入配置页签时平滑滚动到顶部
    # 页签切换过渡动画：平滑滚动到顶部
    def animate_to_top():
        steps = 10
        for i in range(steps + 1):
            container.after(i * 15, lambda v=i/steps: mapping_canvas.yview_moveto(1.0 - v))
    try:
        container.winfo_toplevel().bind('<<NotebookTabChanged>>', lambda e: animate_to_top())
    except Exception:
        pass


if __name__ == "__main__":
    create_gui()
