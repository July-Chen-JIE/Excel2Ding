import re
import pandas as pd


def deep_clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_columns = []
    for col in df.columns:
        if str(col).startswith('Unnamed:'):
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
    return df.dropna(how='all')


def dynamic_column_matching(df, column_mapper):
    column_mapping = column_mapper.get_mapping()
    matched = {}
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
    return matched

