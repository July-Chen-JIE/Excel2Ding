import tkinter as tk
from tkinter import ttk
import logging

try:
    from ui_config import LABEL_FONT, TEXT_COLOR, ENTRY_FONT
except ImportError:
    LABEL_FONT = ('Microsoft YaHei UI', 10)
    TEXT_COLOR = "#2D3748"
    ENTRY_FONT = ('Microsoft YaHei UI', 9)


class ProductLineManager:
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
            self.frame, text="删除", style='Danger.TButton',
            command=lambda i=idx: self.remove_row(i)
        )
        delete_btn.grid(row=idx, column=4, padx=(5, 0), pady=(8, 10))

        self.rows.append({
            'product_var': product_var,
            'contact_var': contact_var,
            'product_entry': product_entry,
            'contact_entry': contact_entry,
            'delete_btn': delete_btn
        })

    def remove_row(self, idx):
        if not (0 <= idx < len(self.rows)):
            return
        for widget in self.frame.grid_slaves(row=idx):
            widget.destroy()
        self.rows.pop(idx)
        self._rebuild_grid()

    def _rebuild_grid(self):
        for i, row_data in enumerate(self.rows):
            for widget in self.frame.grid_slaves(row=i + 1):
                widget.grid(row=i)
            row_data['delete_btn'].config(command=lambda j=i: self.remove_row(j))

    def clear(self):
        while self.rows:
            self.remove_row(0)

    def get_mappings(self):
        mappings = []
        for row_data in self.rows:
            p = row_data['product_var'].get().strip()
            c = row_data['contact_var'].get().strip()
            if p and c:
                mappings.append((p, c))
        return mappings

    def load_from_file(self, path='product_mapping.json'):
        if not path:
            path = 'product_mapping.json'
        try:
            if not path or not path.startswith('/') and not path.startswith('\\') and ':' not in path:
                if not path.startswith('./') and not path.startswith('.\\'):
                    path = path
                else:
                    import os
                    path = os.path.abspath(path)
        except Exception:
            pass
        try:
            import os
            if os.path.exists(path):
                import json
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for item in data.get('mappings', []):
                    self.add_row(item.get('product', ''), item.get('contact', ''))
        except Exception as e:
            logging.warning("加载产品线映射失败: %s", e)

    def save_to_file(self, path='product_mapping.json'):
        if not path:
            path = 'product_mapping.json'
        try:
            import json
            mappings_data = [{'product': p, 'contact': c} for p, c in self.get_mappings()]
            with open(path, 'w', encoding='utf-8') as f:
                json.dump({'mappings': mappings_data}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.warning("保存产品线映射失败: %s", e)
