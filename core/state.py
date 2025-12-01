from dataclasses import dataclass


@dataclass
class AppState:
    input_file: str = ""
    output_dir: str = ""
    start_date: str = ""
    end_date: str = ""
