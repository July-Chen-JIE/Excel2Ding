from dataclasses import dataclass


@dataclass
class AppState:
    input_file: str = ""
    output_dir: str = ""
    start_date: str = ""
    end_date: str = ""
    # 预留其他可选设置字段
    # e.g. enable_filters: bool = False
