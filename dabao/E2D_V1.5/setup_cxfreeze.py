#!/usr/bin/env python3
"""cx_Freeze配置脚本，用于创建Excel2Ding Windows安装程序"""

from cx_Freeze import setup, Executable
import sys
import os
import shutil

# 确保中文显示正常
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 创建临时文件夹用于存放column_mapping.json
temp_dir = os.path.join(current_dir, "temp")
os.makedirs(temp_dir, exist_ok=True)

# 复制column_mapping.json文件到临时目录
column_mapping_source = os.path.join(current_dir, "../Excel2Ding_V1.3/column_mapping.json")
column_mapping_target = os.path.join(temp_dir, "column_mapping.json")
if os.path.exists(column_mapping_source):
    shutil.copy2(column_mapping_source, column_mapping_target)
else:
    # 如果源文件不存在，创建一个默认的column_mapping.json
    default_mapping = '''{
  "mapping": {
    "发起人姓名": ["发起人姓名", "对接人"],
    "发起时间": ["发起时间", "创建时间"],
    "当前周": ["当前周"],
    "项目名称": ["项目名称"],
    "产品线": ["产品线", "产品"],
    "申请状态": ["申请状态", "当前进度"],
    "特制化比例": ["特制化比例(%)", "特制化比例"],
    "可常规化比例": ["可常规化比例(%)", "可常规化比例"],
    "建议报价元": ["建议报价(元)", "报价金额"],
    "定制内容": ["定制内容"],
    "软件版本": ["软件版本/产品名称", "产品名称"],
    "硬件情况": ["硬件情况（分辨率）/原产品主型号", "原产品主型号"],
    "销售部门": ["销售部门"],
    "定制人": ["定制人/销售经理", "销售经理"]
  },
  "output_columns": {
    "发起人姓名": "对接人（发起人）",
    "发起时间": "发起时间",
    "当前周": "当前周",
    "项目名称": "项目名称",
    "产品线": "产品线",
    "申请状态": "当前进度",
    "特制化比例": "特制化比例(%)",
    "可常规化比例": "可常规化比例(%)",
    "建议报价元": "建议报价(元)",
    "定制内容": "定制内容",
    "软件版本": "软件版本/产品名称",
    "硬件情况": "硬件情况（分辨率）/原产品主型号",
    "销售部门": "销售部门",
    "定制人": "定制人/销售经理"
  }
}'''
    with open(column_mapping_target, 'w', encoding='utf-8') as f:
        f.write(default_mapping)

# 设置打包选项
build_exe_options = {
    "packages": [
        "tkinter", 
        "pandas",
        "openpyxl",
        "tkcalendar",
        "datetime",
        "json",
        "re",
        "traceback",
        "warnings",
        "os"
    ],
    "include_files": [
        (os.path.join(current_dir, "../Excel2Ding.ico"), "Excel2Ding.ico"),
        (column_mapping_target, "column_mapping.json")
    ],
    "excludes": [
        "PyQt6",
        "PyQt6.QtCore", 
        "PyQt6.QtGui", 
        "PyQt6.QtWidgets",
        "PyQt6.QtWebEngineWidgets",
        "PyQt6.QtWebEngineCore",
        "PyPDF2",
        "ebooklib",
        "test",
        "unittest",
        "setuptools",
        "pip",
        "jedi",  # 排除jedi库
        "django",  # 排除django相关模块
        "scipy",
        "matplotlib",
        "requests",
        "bs4",
        "urllib3",
        "PIL"
    ],
    "include_msvcr": True  # 包含Microsoft Visual C++运行时
}

# 设置安装程序选项
bdist_msi_options = {
    "add_to_path": False,
    "all_users": True,
    "initial_target_dir": "C:\\Program Files\\Excel2Ding",
    "install_icon": os.path.join(current_dir, "../Excel2Ding.ico"),
    "summary_data": {
        "author": "Excel2DingStudio",
        "comments": "Excel数据处理工具",
        "keywords": "Excel,数据处理,自动化"
    },
    # 配置卸载程序相关选项
    "upgrade_code": "{87654321-4321-5678-90AB-CDEF01234567}",  # 有效的GUID格式升级码
    "product_code": "{12345678-8765-4321-90AB-CDEF01234567}"  # 有效的GUID格式产品码
}

# 定义可执行文件
base = "Win32GUI" if sys.platform == "win32" else None

executables = [
    Executable(
        script=os.path.join(current_dir, "../Excel2Ding_1.5.py"),
        base=base,
        target_name="Excel2Ding.exe",
        icon=os.path.join(current_dir, "../Excel2Ding.ico"),
        shortcut_name="Excel2Ding",
        shortcut_dir="DesktopFolder"
    )
]

# 运行setup函数
setup(
    name="Excel2Ding",
    version="1.5",  
    description="Excel数据处理工具",
    author="Excel2DingStudio",
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options
    },
    executables=executables
)