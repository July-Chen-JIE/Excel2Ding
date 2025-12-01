import json
import os


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
        try:
            if os.path.exists('column_mapping.json'):
                with open('column_mapping.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.column_mapping = data.get('mapping', self.DEFAULT_MAPPING)
                    self.output_columns = data.get('output_columns', self.OUTPUT_COLUMNS)
            else:
                self.column_mapping = self.DEFAULT_MAPPING
                self.output_columns = self.OUTPUT_COLUMNS
        except Exception:
            self.column_mapping = self.DEFAULT_MAPPING
            self.output_columns = self.OUTPUT_COLUMNS

    def save_mapping(self):
        try:
            with open('column_mapping.json', 'w', encoding='utf-8') as f:
                json.dump({'mapping': self.column_mapping, 'output_columns': self.output_columns}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_from_path(self, path: str):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.column_mapping = data.get('mapping', self.DEFAULT_MAPPING)
                self.output_columns = data.get('output_columns', self.OUTPUT_COLUMNS)
        except Exception:
            pass

    def save_to_path(self, path: str):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump({'mapping': self.column_mapping, 'output_columns': self.output_columns}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get_mapping(self):
        return self.column_mapping

    def get_output_columns(self):
        return self.output_columns

