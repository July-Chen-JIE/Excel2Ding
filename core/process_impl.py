import os
import re
import pandas as pd
from openpyxl.styles import Alignment
from core import mapping as mapping_core
from core import transform as transform_core


def get_sheets_with_data(file_path):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheets_with_data = []
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=10)
                if not df.empty and len(df) > 0:
                    first_row = df.iloc[0].astype(str)
                    non_empty_count = first_row.count()
                    if non_empty_count >= 5:
                        header_keywords = ['时间', '日期', '申请', '审批', '金额', '报价', '产品', '类型']
                        first_row_text = ' '.join(first_row.tolist()).lower()
                        if any(keyword in first_row_text for keyword in header_keywords):
                            sheets_with_data.append(sheet_name)
                        elif len(df.columns) >= 10:
                            sheets_with_data.append(sheet_name)
            except Exception:
                continue
        return sheets_with_data
    except Exception:
        return []


def process_raw_excel(input_file, output_file, start_date=None, end_date=None, target_product=None, new_contact=None, product_contact_list=None, replace_mode='overwrite', progress_callback=None, cancel_event=None):
    try:
        if progress_callback:
            progress_callback(10, "正在分析文件结构...")
        sheet_names = get_sheets_with_data(input_file)
        if not sheet_names:
            raise Exception("未找到包含数据的工作表")
        if progress_callback:
            progress_callback(20, f"发现 {len(sheet_names)} 个工作表: {sheet_names}")
        all_data = []
        for i, sheet_name in enumerate(sheet_names):
            if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
                return False
            try:
                if progress_callback:
                    progress_callback(20 + i * 20 // len(sheet_names), f"正在读取工作表: {sheet_name}")
                df = pd.read_excel(input_file, sheet_name=sheet_name, header=1, converters={'发起时间': str})
                df = transform_core.deep_clean_columns(df)
                df['数据来源'] = sheet_name
                all_data.append(df)
            except Exception:
                continue
        if not all_data:
            raise Exception("未能读取任何工作表数据")
        if progress_callback:
            progress_callback(40, "合并所有工作表数据...")
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
        combined_df = pd.concat(all_data, ignore_index=True)
        if progress_callback:
            progress_callback(50, f"数据合并完成，共 {len(combined_df)} 行记录")
        if progress_callback:
            progress_callback(60, "正在匹配列名...")
        column_mapper = mapping_core.ColumnMapper()
        matched = transform_core.dynamic_column_matching(combined_df, column_mapper)
        if start_date and end_date:
            if progress_callback:
                progress_callback(70, f"筛选日期范围: {start_date} 至 {end_date}")
            try:
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    combined_df['parsed_time'] = pd.to_datetime(combined_df[time_column].astype(str), errors='coerce')
                    if combined_df['parsed_time'].isna().all():
                        date_pattern = r'(\d{4}-\d{2}-\d{2})'
                        def _extract_ymd(text):
                            m = re.search(date_pattern, str(text))
                            return m.group(1) if m else None
                        combined_df['parsed_time'] = pd.to_datetime(combined_df[time_column].map(_extract_ymd), errors='coerce')
                else:
                    combined_df['parsed_time'] = pd.to_datetime(combined_df.get('发起时间', pd.Series([pd.NaT] * len(combined_df))).astype(str), errors='coerce')
                if combined_df['parsed_time'].isna().all():
                    date_any_pattern = re.compile(r'(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?)')
                    vals = []
                    for _, row in combined_df.iterrows():
                        text_line = ' '.join([str(v) for v in row.values])
                        m = date_any_pattern.search(text_line)
                        vals.append(m.group(1) if m else None)
                    combined_df['parsed_time'] = pd.to_datetime(pd.Series(vals), errors='coerce')
                mask = (combined_df['parsed_time'] >= start_date) & (combined_df['parsed_time'] <= end_date)
                filtered_df = combined_df[mask]
                if progress_callback:
                    progress_callback(80, f"日期筛选完成，剩余 {len(filtered_df)} 行记录")
            except Exception:
                filtered_df = combined_df
        else:
            if 'parsed_time' not in combined_df.columns:
                time_columns = [col for col in combined_df.columns if '发起时间' in str(col)]
                if time_columns:
                    time_column = time_columns[0]
                    combined_df['parsed_time'] = pd.to_datetime(combined_df[time_column].astype(str), errors='coerce')
                else:
                    combined_df['parsed_time'] = pd.Series([pd.NaT] * len(combined_df))
            filtered_df = combined_df
        if progress_callback:
            progress_callback(90, "正在生成输出数据...")
        filtered_df.loc[:, '当前周'] = filtered_df['parsed_time'].dt.isocalendar().week
        desired_order = [
            '对接人（发起人）','发起时间','当前周','项目名称','产品线','当前进度','特制化比例(%)','可常规化比例(%)','建议报价(元)','定制内容','软件版本/产品名称','硬件情况（分辨率）/原产品主型号','销售部门','定制人/销售经理'
        ]
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
        output_df = pd.DataFrame()
        cm = column_mapper.get_output_columns()
        rev_cm = {v: k for k, v in cm.items()}
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
        for out_col in desired_order:
            filled = False
            if out_col in rev_cm:
                norm = rev_cm[out_col]
                if norm in matched and matched[norm] in filtered_df.columns:
                    output_df[out_col] = filtered_df[matched[norm]]
                    filled = True
            if not filled:
                src = find_source_column(alias_mappings.get(out_col, []))
                if src:
                    output_df[out_col] = filtered_df[src]
                    filled = True
            if not filled and out_col == '当前周':
                output_df[out_col] = filtered_df['parsed_time'].dt.isocalendar().week
                filled = True
            if not filled:
                output_df[out_col] = ""
        try:
            if ('发起时间' in output_df.columns) and (output_df['发起时间'].isna().all() or (output_df['发起时间'] == "").all()):
                output_df['发起时间'] = filtered_df['parsed_time']
        except Exception:
            pass
        if product_contact_list and isinstance(product_contact_list, list):
            if '产品线' in output_df.columns:
                prod_series = output_df['产品线'].astype(str).str.strip()
                all_empty = (prod_series == "").all()
                if all_empty and len(product_contact_list) == 1:
                    default_product, _default_contact = product_contact_list[0]
                    output_df['产品线'] = default_product
            if '产品线' in output_df.columns and '对接人（发起人）' in output_df.columns:
                for product, contact in product_contact_list:
                    mask = output_df['产品线'].astype(str).str.strip().str.lower() == str(product).strip().lower()
                    if str(replace_mode).lower() == 'fill_empty':
                        empty_mask = output_df['对接人（发起人）'].astype(str).str.strip() == ""
                        output_df.loc[mask & empty_mask, '对接人（发起人）'] = contact
                    else:
                        output_df.loc[mask, '对接人（发起人）'] = contact
        elif target_product and new_contact:
            if '产品线' in output_df.columns and '对接人（发起人）' in output_df.columns:
                output_df.loc[output_df['产品线'] == target_product, '对接人（发起人）'] = new_contact
        if '发起时间' in output_df.columns:
            output_df['发起时间'] = pd.to_datetime(output_df['发起时间'], errors='coerce')
            output_df = output_df.sort_values(by='发起时间', ascending=False, na_position='last')
        elif 'parsed_time' in filtered_df.columns:
            output_df = output_df.iloc[filtered_df['parsed_time'].sort_values(ascending=False, na_position='last').index]
        if progress_callback:
            progress_callback(95, f"正在保存结果到: {output_file}")
        if cancel_event and getattr(cancel_event, 'is_set', None) and cancel_event.is_set():
            return False
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='处理结果')
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
                for cell in column:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        if progress_callback:
            progress_callback(100, "文件处理完成!")
        return True
    except Exception as e:
        raise e

