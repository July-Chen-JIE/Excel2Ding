#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""UI 配置常量模块

集中管理颜色、字体和布局常量，供 GUI 使用。
"""

# 窗口与布局 - 增大窗口尺寸以提供更好的空间
WINDOW_SIZE = "610x830"
PADDING = 25

# 颜色方案 - 优化颜色搭配，提升视觉体验
PRIMARY_COLOR = "#409EFF"
SECONDARY_COLOR = "#67C23A"
SUCCESS_COLOR = "#67C23A"
WARNING_COLOR = "#E6A23C"
DANGER_COLOR = "#F56C6C"
INFO_COLOR = "#409EFF"

# 背景色系
BG_COLOR = "#F5F7FA"
PANEL_BG = "#FFFFFF"              # 面板白色背景
CARD_BG = "#FFFFFF"               # 卡片白色背景
HOVER_BG = "#F8FAFC"              # 悬停背景

# 文字色系
TEXT_COLOR = "#2D3748"
SECONDARY_TEXT = "#606266"
PLACEHOLDER_COLOR = "#A8ABB2"

# 边框和分割线
BORDER_COLOR = "#DCDFE6"
LIGHT_BORDER = "#EBEEF5"
FOCUS_BORDER = "#409EFF"

# 阴影效果
SHADOW_COLOR = "#E9EDF3"

# 字体配置 - 优化字体大小，提升可读性
TITLE_FONT = ('Microsoft YaHei UI', 14, 'bold')         # 主标题
SUBTITLE_FONT = ('Microsoft YaHei UI', 12, 'bold')      # 副标题
LABEL_FONT = ('Microsoft YaHei UI', 10)                 # 标签文字
BUTTON_FONT = ('Microsoft YaHei UI', 9, 'bold')         # 按钮文字
ENTRY_FONT = ('Microsoft YaHei UI', 9)                  # 输入框文字
SMALL_FONT = ('Microsoft YaHei UI', 8)                  # 小号文字

# 间距配置
BUTTON_PADDING = (12, 10)
ENTRY_PADDING = (8, 6)
CARD_PADDING = 16
GROUP_PADDING = 12

def apply_design_system(style):
    style.configure('TButton', padding=BUTTON_PADDING, relief='flat', background=PRIMARY_COLOR, foreground='white', font=BUTTON_FONT, borderwidth=0)
    style.map('TButton', background=[('active', PRIMARY_COLOR), ('pressed', '#1E3A8A'), ('disabled', '#94A3B8')], foreground=[('disabled', '#CBD5E1')])
    style.configure('Primary.TButton', padding=BUTTON_PADDING, relief='flat', background=PRIMARY_COLOR, foreground='white', font=BUTTON_FONT, borderwidth=0)
    style.configure('Secondary.TButton', padding=BUTTON_PADDING, relief='flat', background=HOVER_BG, foreground=SECONDARY_TEXT, font=BUTTON_FONT, borderwidth=0)
    style.map('Secondary.TButton', background=[('active', '#E2E8F0'), ('pressed', '#CBD5E1'), ('disabled', '#F8FAFC')], foreground=[('disabled', '#94A3B8')])
    style.configure('Danger.TButton', padding=BUTTON_PADDING, relief='flat', background=DANGER_COLOR, foreground='white', font=BUTTON_FONT, borderwidth=0)
    style.map('Danger.TButton', background=[('active', '#DC2626'), ('pressed', '#B91C1C'), ('disabled', '#FCA5A5')])
    style.configure('Info.TButton', padding=BUTTON_PADDING, relief='flat', background=INFO_COLOR, foreground='white', font=BUTTON_FONT, borderwidth=0)
    style.configure('TLabel', font=LABEL_FONT, foreground=TEXT_COLOR, padding=(0, 8))
    style.configure('Card.TLabel', background=CARD_BG, font=LABEL_FONT, foreground=TEXT_COLOR, padding=(0, 8))
    style.configure('TEntry', padding=ENTRY_PADDING, relief='flat', borderwidth=2, font=ENTRY_FONT, foreground=TEXT_COLOR, fieldbackground='white')
    style.map('TEntry', foreground=[('disabled', '#9CA3AF')])
    style.configure('TFrame', background=BG_COLOR, borderwidth=0)
    style.configure('Card.TFrame', background=CARD_BG, borderwidth=0)
    style.configure('TLabelframe', background=CARD_BG, borderwidth=0, relief='flat', padding=(CARD_PADDING, CARD_PADDING))
    style.configure('TLabelframe.Label', background=CARD_BG, font=SUBTITLE_FONT, foreground=TEXT_COLOR, padding=(0, 0, 0, 12))
    style.configure('Card.TLabelframe', background=CARD_BG, borderwidth=0, relief='flat', padding=(CARD_PADDING, CARD_PADDING))
    style.configure('Card.TLabelframe.Label', background=CARD_BG, font=SUBTITLE_FONT, foreground=TEXT_COLOR, padding=(0, 0, 0, 12))
    style.configure('Vertical.TScrollbar', gripcount=0, background=LIGHT_BORDER, darkcolor=LIGHT_BORDER, lightcolor=LIGHT_BORDER, troughcolor=CARD_BG, bordercolor=LIGHT_BORDER, arrowcolor=SECONDARY_TEXT)
    style.configure('Horizontal.TScrollbar', gripcount=0, background=LIGHT_BORDER, darkcolor=LIGHT_BORDER, lightcolor=LIGHT_BORDER, troughcolor=CARD_BG, bordercolor=LIGHT_BORDER, arrowcolor=SECONDARY_TEXT)
    style.configure('TNotebook', background=BG_COLOR, borderwidth=0, tabmargins=(6, 6, 6, 0))
    style.configure('TNotebook.Tab', padding=(6, 4), font=('Microsoft YaHei UI', 11), foreground=SECONDARY_TEXT)
    style.map('TNotebook.Tab',
              background=[('selected', 'white'), ('active', '#F2F6FC')],
              foreground=[('selected', TEXT_COLOR), ('active', TEXT_COLOR)])
    style.configure('primary.TButton', padding=BUTTON_PADDING, font=BUTTON_FONT)
    style.configure('danger.TButton', padding=BUTTON_PADDING, font=BUTTON_FONT)
    style.configure('info.TButton', padding=BUTTON_PADDING, font=BUTTON_FONT)
