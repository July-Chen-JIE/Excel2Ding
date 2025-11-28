import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import re
import traceback
from tkinter import ttk
from tkcalendar import DateEntry
import os
import json

# 将配色方案移到文件顶部
PRIMARY_COLOR = "#409EFF"      # 主色调（科技蓝）
SECONDARY_COLOR = "#67C23A"    # 辅助色（绿色）
BG_COLOR = "#F5F7FA"          # 背景色（浅灰）
TEXT_COLOR = "#2D3748"        # 主文本色（深灰）
BUTTON_TEXT_COLOR = "white"   # 按钮文字颜色
ERROR_COLOR = "#F56C6C"       # 错误提示色（红色）
BORDER_COLOR = "#DCDFE6"      # 边框颜色（浅灰）

# 窗口布局常量
WINDOW_PADDING = 25
SECTION_SPACING = 15
WIDGET_SPACING = 10
BUTTON_SPACING = 5

# 窗口大小设置
MAIN_WINDOW_SIZE = "600x800"
CONFIG_WINDOW_SIZE = "800x550"
EDIT_WINDOW_SIZE = "400x350"
PROGRESS_WINDOW_SIZE = "450x180"

class ColumnMapper:
    """列映射管理类
    
    负责管理Excel列名的映射关系和输出配置。提供配置的加载、保存和获取功能。
    
    Attributes:
        DEFAULT_MAPPING (dict): 默认的列名映射配置
        OUTPUT_COLUMNS (dict): 默认的输出列名配置
        column_mapping (dict): 当前使用的列名映射
        output_columns (dict): 当前使用的输出列名
    """
    
    DEFAULT_MAPPING = {
        '发起人姓名': ['发起人姓名', '姓名'],
        '发起时间': ['发起时间', '创建时间'],
        '项目名称': ['项目名称', '项目'],
        '产品线': ['产品线', '产品'],
        '建议报价元': ['建议报价(元)', '报价金额'],
        '申请状态': ['申请状态', '当前进度']
    }
    
    OUTPUT_COLUMNS = {
        '发起人姓名': '对接人',
        '发起时间': '创建时间',
        '当前周': '当前周',
        '项目名称': '项目名称',
        '产品线': '产品',
        '申请状态': '当前进度',
        '建议报价元': '报价金额'
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
            raise ValueError("日期解析失败，请检查“发起时间”列是否为有效的 Excel 序列号格式")
        
        # 时间范围过滤
        start_dt = datetime.strptime(start_date, "%Y/%m/%d")
        end_dt = datetime.strptime(end_date, "%Y/%m/%d")
        mask = (valid_df['datetime_obj'].dt.date >= start_dt.date()) & \
               (valid_df['datetime_obj'].dt.date <= end_dt.date())
        filtered = valid_df[mask].copy()  # 创建副本避免警告
        
        # 检查过滤后的数据是否为空
        if filtered.empty:
            raise ValueError("所选时间范围内没有数据！")
        
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
                
                # 添加安全检查
                if not output_df.empty:
                    try:
                        max_length = max(
                            max_length,
                            len(str(col)) * 2,
                            max((len(str(cell)) * 1.2 for cell in output_df[col].astype(str) if pd.notna(cell)), default=0)
                        )
                    except Exception as e:
                        print(f"警告：计算列 {col} 宽度时出错：{str(e)}")
                        max_length = 20  # 设置默认宽度
                else:
                    max_length = 20  # 设置默认宽度
                
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
        # 修改为绝对路径，确保图标文件存在
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Excel2Ding.ico")
        if os.path.exists(icon_path):
            window.iconbitmap(icon_path)
        else:
            print(f"图标文件不存在: {icon_path}")
    except Exception as e:
        print(f"加载图标失败: {e}")

def setup_window(window, title, size, resizable=(False, False)):
    """统一设置窗口属性"""
    window.title(title)
    window.geometry(size)
    window.configure(bg=BG_COLOR)
    window.resizable(*resizable)
    set_window_icon(window)
    center_window(window)

def create_mapping_window(root: tk.Tk) -> None:
    """创建列映射配置窗口
    
    创建一个模态配置窗口，用于管理列名映射关系。
    
    Args:
        root: 主窗口实例
    """
    config_window = tk.Toplevel(root)
    setup_window(config_window, "列映射配置", CONFIG_WINDOW_SIZE)
    config_window.transient(root)
    config_window.grab_set()
    
    # 创建主框架
    main_frame = ttk.Frame(config_window, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 添加说明文本
    tips_frame = ttk.LabelFrame(main_frame, text="▌配置说明")
    tips_frame.pack(fill=tk.X, pady=(0, 10))
    
    tips_text = """配置说明：
• 目标列名：程序内部使用的标准列名
• 映射别名：Excel中可能出现的列名（多个用英文逗号","分隔）
• 输出列名：最终输出Excel文件中显示的列名
注意：映射别名必须使用英文逗号","分隔，不能使用中文逗号"，"！"""
    
    ttk.Label(tips_frame, text=tips_text, justify=tk.LEFT,
              font=('Microsoft YaHei UI', 10)).pack(anchor="w", pady=5)
    
    # 创建列表框架
    list_frame = ttk.Frame(main_frame)
    list_frame.pack(fill=tk.BOTH, expand=True)
    
    # 创建映射列表
    columns = ('目标列名', '映射别名', '输出列名')
    tree = ttk.Treeview(list_frame, columns=columns, show='headings')
    
    # 设置列标题
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200)
    
    # 创建滚动条
    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    # 放置列表和滚动条
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # 加载当前配置
    mapper = ColumnMapper()
    mapping = mapper.column_mapping
    output_cols = mapper.output_columns
    
    def load_mapping():
        """加载映射到列表"""
        tree.delete(*tree.get_children())  # 清空列表
        for target, aliases in mapping.items():
            output_name = output_cols.get(target, target)
            tree.insert('', tk.END, values=(target, ', '.join(aliases), output_name))
    
    def save_mapping():
        """保存映射配置"""
        try:
            new_mapping = {}
            new_output_cols = {}
            for item in tree.get_children():
                values = tree.item(item)['values']
                target = values[0]
                aliases = [alias.strip() for alias in values[1].split(',')]
                output_name = values[2]
                new_mapping[target] = aliases
                new_output_cols[target] = output_name
            
            mapper.column_mapping = new_mapping
            mapper.output_columns = new_output_cols
            mapper.save_mapping()
            messagebox.showinfo("成功", "配置已保存！")
            config_window.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def edit_item():
        """编辑选中项"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个配置项！")
            return
        
        item = tree.item(selected[0])
        values = item['values']
        
        # 创建编辑窗口
        edit_window = tk.Toplevel(config_window)
        setup_window(edit_window, "编辑映射", EDIT_WINDOW_SIZE)
        edit_window.transient(config_window)
        
        main_frame = ttk.Frame(edit_window, padding=20)  # 添加主框架
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(edit_window, text="目标列名:").pack(pady=5)
        target_entry = ttk.Entry(edit_window)
        target_entry.insert(0, values[0])
        target_entry.pack(fill=tk.X, padx=20)
        
        ttk.Label(edit_window, text="映射别名 (用逗号分隔):").pack(pady=5)
        aliases_entry = ttk.Entry(edit_window)
        aliases_entry.insert(0, values[1])
        aliases_entry.pack(fill=tk.X, padx=20)
        
        ttk.Label(edit_window, text="输出列名:").pack(pady=5)
        output_entry = ttk.Entry(edit_window)
        output_entry.insert(0, values[2])
        output_entry.pack(fill=tk.X, padx=20)
        
        def update():
            """更新列表项"""
            tree.item(selected[0], values=(
                target_entry.get(),
                aliases_entry.get(),
                output_entry.get()
            ))
            edit_window.destroy()
        
        ttk.Button(edit_window, text="确定", 
               command=update, 
               style='Modern.TButton').pack(pady=20)  # 修改这里
    
    # 添加增加和删除按钮
    def add_item():
        """添加新配置项"""
        edit_window = tk.Toplevel(config_window)
        setup_window(edit_window, "添加映射", EDIT_WINDOW_SIZE)
        edit_window.transient(config_window)
        
        main_frame = ttk.Frame(edit_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="目标列名:").pack(pady=5)
        target_entry = ttk.Entry(main_frame)
        target_entry.pack(fill=tk.X, padx=20)
        
        ttk.Label(main_frame, text="映射别名 (用英文逗号\",\"分隔):").pack(pady=5)
        aliases_entry = ttk.Entry(main_frame)
        aliases_entry.pack(fill=tk.X, padx=20)
        
        ttk.Label(main_frame, text="输出列名:").pack(pady=5)
        output_entry = ttk.Entry(main_frame)
        output_entry.pack(fill=tk.X, padx=20)
        
        def insert():
            """插入新项"""
            tree.insert('', tk.END, values=(
                target_entry.get(),
                aliases_entry.get(),
                output_entry.get()
            ))
            edit_window.destroy()
        
        ttk.Button(main_frame, text="确定", command=insert,
                  style='Modern.TButton').pack(pady=20)
    
    def delete_item():
        """删除选中项"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的配置项！")
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的配置项吗？"):
            for item in selected:
                tree.delete(item)
    
    # 修改按钮框架
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(20, 0))
    
    ttk.Button(button_frame, text="➕ 添加", command=add_item,
               style='Modern.TButton').pack(side=tk.LEFT)
    ttk.Button(button_frame, text="✏️ 编辑", command=edit_item,
               style='Modern.TButton').pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="❌ 删除", command=delete_item,
               style='Modern.TButton').pack(side=tk.LEFT)
    ttk.Button(button_frame, text="💾 保存", command=save_mapping,
               style='Modern.TButton').pack(side=tk.RIGHT)
    ttk.Button(button_frame, text="取消", command=config_window.destroy,
               style='Modern.TButton').pack(side=tk.RIGHT, padx=5)
    
    # 加载当前配置
    load_mapping()
    
    # 设置模态窗口
    config_window.grab_set()
    config_window.focus_set()
    center_window(config_window)

def setup_styles(style: ttk.Style):
    """设置应用程序统一样式"""
    style.theme_use('clam')
    
    # 基础样式配置
    PADDING = {
        'button': (20, 10),
        'entry': (10, 8),
        'frame': 15,
        'labelframe': 20,
        'treeview': 10
    }
    
    # 基础框架样式
    style.configure('TFrame',
        background=BG_COLOR)
    
    # 标签样式
    style.configure('TLabel',
        font=('Microsoft YaHei UI', 10),
        padding=8,
        background=BG_COLOR)
    
    # 标签框样式
    style.configure('TLabelframe',
        background=BG_COLOR,
        padding=PADDING['labelframe'])
    
    style.configure('TLabelframe.Label',
        font=('Microsoft YaHei UI', 11, 'bold'),
        foreground=PRIMARY_COLOR,
        background=BG_COLOR)
    
    # 按钮样式
    style.configure('Modern.TButton',
        font=('Microsoft YaHei UI', 10, 'bold'),
        padding=PADDING['button'],
        background=PRIMARY_COLOR,
        foreground=BUTTON_TEXT_COLOR,
        borderwidth=0,
        relief="flat")
    
    style.map('Modern.TButton',
        background=[('pressed', '#3a8ee6'), ('active', '#79BBFF')],
        foreground=[('pressed', 'white'), ('active', 'white')])
    
    # 输入框样式
    style.configure('TEntry',
        font=('Microsoft YaHei UI', 10),
        padding=PADDING['entry'],
        fieldbackground='white',
        borderwidth=1,
        relief="solid")
    
    # 树状视图样式
    style.configure('Treeview',
        background='white',
        fieldbackground='white',
        font=('Microsoft YaHei UI', 10),
        rowheight=35,
        padding=PADDING['treeview'])
    
    style.configure('Treeview.Heading',
        font=('Microsoft YaHei UI', 10, 'bold'),
        padding=8,
        background=BG_COLOR,
        foreground=TEXT_COLOR)
    
    # 进度条样式
    style.configure('Modern.Horizontal.TProgressbar',
        troughcolor='#F3F4F6',
        background=PRIMARY_COLOR,
        thickness=10,
        borderwidth=0,
        relief="flat")

def create_dialog_frame(window: tk.Toplevel, title: str) -> ttk.Frame:
    """创建统一的对话框框架"""
    window.configure(bg=BG_COLOR)
    window.title(title)
    
    # 主框架
    main_frame = ttk.Frame(window, padding=WINDOW_PADDING)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    return main_frame

def create_button_container(parent: ttk.Frame, padding: int = BUTTON_SPACING) -> ttk.Frame:
    """创建统一的按钮容器"""
    container = ttk.Frame(parent)
    container.pack(fill=tk.X, pady=(padding, 0))
    return container

def create_info_label(parent: ttk.Frame, text: str, is_warning: bool = False) -> ttk.Label:
    """创建统一的信息标签"""
    return ttk.Label(
        parent,
        text=text,
        justify=tk.LEFT,
        font=('Microsoft YaHei UI', 10),
        foreground=ERROR_COLOR if is_warning else TEXT_COLOR,
        background=BG_COLOR,
        wraplength=600
    )

# 在 create_gui 函数开始处添加日期选择器样式
def create_gui():
    root = tk.Tk()
    setup_window(root, "定制审批单处理工具", MAIN_WINDOW_SIZE)
    style = ttk.Style()
    setup_styles(style)
    
    # 声明变量
    global PRIMARY_COLOR, SECONDARY_COLOR, BG_COLOR, TEXT_COLOR, BUTTON_TEXT_COLOR, ERROR_COLOR, BORDER_COLOR
    global start_date, end_date, input_entry, output_entry, target_product, new_contact
    global process_btn, exit_btn
    
    # 日期选择器样式
    date_style = {
        'font': ('Microsoft YaHei UI', 10),
        'background': 'white',
        'foreground': TEXT_COLOR,
        'borderwidth': 1,
        'width': 12,
        'relief': "solid",
        'date_pattern': 'y/mm/dd',  # 添加这行，指定日期格式
        'locale': 'zh_CN'  # 添加这行，指定中文区域
    }
    
    # 设置窗口
    root.title("定制审批单处理工具")
    root.geometry(MAIN_WINDOW_SIZE)
    root.configure(bg=BG_COLOR)
    root.resizable(False, False)  # 禁止调整大小以保持布局一致性
    
    # 创建主框架
    main_frame = ttk.Frame(root, padding=WINDOW_PADDING)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
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
        target_prod = target_product.get().strip()
        new_cont = new_contact.get().strip()
        
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
        
        # 创建进度条窗口
        progress_window, progress_var, progress_label = create_progress_window(root)
        
        try:
            if process_excel(
                input_file,
                start_date.get(),
                end_date.get(),
                f"{output_dir}/处理结果_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                target_product=target_prod if target_prod else None,
                new_contact=new_cont if new_cont else None,
                progress_callback=lambda p, msg: update_progress(progress_var, progress_label, p, msg)
            ):
                messagebox.showinfo("完成", "文件处理成功！")
        finally:
            # 关闭进度条窗口
            progress_window.destroy()
            # 恢复按钮状态
            process_btn.configure(state='normal')
            exit_btn.configure(state='normal')

    # # 更新配色方案，使用经典配色
    # PRIMARY_COLOR = "#409EFF"      # 主色调（科技蓝）
    # SECONDARY_COLOR = "#67C23A"    # 辅助色（绿色）
    # BG_COLOR = "#F5F7FA"          # 背景色（浅灰）
    # TEXT_COLOR = "#2D3748"        # 主文本色（深灰）
    # BUTTON_TEXT_COLOR = "white"   # 按钮文字颜色
    # ERROR_COLOR = "#F56C6C"       # 错误提示色（红色）
    # BORDER_COLOR = "#DCDFE6"       # 边框颜色（浅灰）

    # 修改控件样式配置
    style = ttk.Style()
    setup_styles(style)

    # 修改窗口基础设置
    root.title("定制审批单处理工具")
    root.geometry(MAIN_WINDOW_SIZE)  # 适当调整窗口大小
    root.configure(bg=BG_COLOR)
    root.resizable(True, True)

    # 调整各区域的间距和内边距
    main_frame = ttk.Frame(root, padding=25)  # 增加主框架内边距
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 各区域之间的间距
    SECTION_PADDING = 15  # 区域间距

    # 输入框和按钮的统一样式
    ENTRY_STYLE = {'font': ('Segoe UI', 10), 'padding': 8}
    BUTTON_STYLE = {'font': ('Segoe UI', 10), 'padding': (15, 8)}

    # 警告文本样式
    WARNING_STYLE = {'font': ('Segoe UI', 10), 'foreground': ERROR_COLOR}

    # 提示框样式
    tips_frame = ttk.LabelFrame(main_frame, text="▌使用提示")
    tips_frame.pack(fill=tk.X, pady=(0, SECTION_PADDING))
    
    tips_text = """⚠️ 使用软件前请手动处理Excel文件:
1、Excel内只能包含一张表格
2、请删除第一行的说明
3、请将【发起时间】单元格格式调整为【文本】"""
    
    ttk.Label(tips_frame, text=tips_text, justify=tk.LEFT,
              font=('Microsoft YaHei UI', 10), foreground="#FF4D4F").pack(anchor="w", pady=5)
    
    # 2. 日期范围设置
    date_frame = ttk.LabelFrame(main_frame, text="▌日期范围设置")
    date_frame.pack(fill=tk.X, pady=10)
    
    date_select_frame = ttk.Frame(date_frame)
    date_select_frame.pack(fill=tk.X, pady=5)
        
    ttk.Label(date_select_frame, text="📅 开始日期:").pack(side=tk.LEFT)
    start_date = DateEntry(date_select_frame, **date_style)
    start_date.pack(side=tk.LEFT, padx=(5, 20))
    
    ttk.Label(date_select_frame, text="📅 结束日期:").pack(side=tk.LEFT)
    # Define date_style before using it

    end_date = DateEntry(date_select_frame, **date_style)
    end_date.pack(side=tk.LEFT, padx=5)
    
    # 日期快捷按钮
    date_buttons = ttk.Frame(date_frame)
    date_buttons.pack(fill=tk.X, pady=(5,0))
    
    ttk.Button(date_buttons, text="📅 本周", command=set_week_start_end,
               style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(date_buttons, text="📆 本月", command=set_month_start_end,
               style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 5))
    ttk.Button(date_buttons, text="🔄 恢复", command=clear_dates,
               style='Modern.TButton').pack(side=tk.LEFT)
    
    # 3. 文件设置（合并输入输出）
    file_frame = ttk.LabelFrame(main_frame, text="▌文件设置")
    file_frame.pack(fill=tk.X, pady=10)
    
    # 输入文件
    input_container = ttk.Frame(file_frame)
    input_container.pack(fill=tk.X, pady=5)
    
    ttk.Label(input_container, text="📂 输入文件:").pack(side=tk.LEFT, padx=(0, 5))
    input_entry = ttk.Entry(input_container, font=('Microsoft YaHei UI', 10))
    input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    
    def select_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, file_path)
            # 自动设置输出目录为输入文件所在目录
            output_entry.delete(0, tk.END)
            output_entry.insert(0, os.path.dirname(file_path))
    
    ttk.Button(input_container, text="浏览", command=select_input_file,
               style='Modern.TButton').pack(side=tk.RIGHT)
    
    # 输出目录
    output_container = ttk.Frame(file_frame)
    output_container.pack(fill=tk.X, pady=5)
    
    ttk.Label(output_container, text="💾 输出目录:").pack(side=tk.LEFT, padx=(0, 5))
    output_entry = ttk.Entry(output_container, font=('Microsoft YaHei UI', 10))
    output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    
    # 修改输出目录浏览按钮的命令
    def select_output_dir():
        dir_path = filedialog.askdirectory()
        if dir_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, dir_path)

    ttk.Button(output_container, text="浏览",
               command=select_output_dir,
               style='Modern.TButton').pack(side=tk.RIGHT)
    
    # 4. 产品线过滤（可选）
    filter_frame = ttk.LabelFrame(main_frame, text="▌产品线过滤设置（可选）")
    filter_frame.pack(fill=tk.X, pady=10)
    
    filter_container = ttk.Frame(filter_frame)
    filter_container.pack(fill=tk.X, pady=5)
    
    ttk.Label(filter_container, text="🔍 目标产品线:").pack(side=tk.LEFT, padx=(0, 5))
    target_product = ttk.Entry(filter_container, width=20, font=('Microsoft YaHei UI', 10))
    target_product.pack(side=tk.LEFT, padx=(0, 20))
    
    ttk.Label(filter_container, text="👤 替换后对接人:").pack(side=tk.LEFT, padx=(0, 5))
    new_contact = ttk.Entry(filter_container, width=20, font=('Microsoft YaHei UI', 10))
    new_contact.pack(side=tk.LEFT)
    
    # 5. 操作按钮
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(15, 0))
    
    process_btn = ttk.Button(button_frame, text="🚀 开始处理",
                            command=start_process, style='Modern.TButton')
    process_btn.pack(side=tk.RIGHT, padx=5)
    
    exit_btn = ttk.Button(button_frame, text="❌ 退出程序",
                         command=root.quit, style='Modern.TButton')
    exit_btn.pack(side=tk.RIGHT)
    
    # 添加配置按钮
    ttk.Button(button_frame, text="⚙️ 配置",
               command=lambda: create_mapping_window(root),
               style='Modern.TButton').pack(side=tk.RIGHT, padx=5)
    
    root.mainloop()

# # 在程序启动时检查图标文件
# def check_resources():
#     """检查必要的资源文件"""
#     icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Excel2Ding.ico")
#     if not os.path.exists(icon_path):
#         print(f"警告: 图标文件不存在 {icon_path}")
#         return False
#     return True

# if __name__ == "__main__":
#     if check_resources():
#         create_gui()
#     else:
#         print("程序资源文件缺失，请确保所有必要文件都存在！")
if __name__ == "__main__":
    create_gui()