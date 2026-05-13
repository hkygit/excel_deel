# services/simplify/output.py
import os
import pandas as pd
from core.excel_writer import write_excel_to_path

def safe_format_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """安全格式化 DataFrame（数值列保留，字符串列清理）"""
    df_copy = df.copy()
    for col in df_copy.columns:
        if pd.api.types.is_numeric_dtype(df_copy[col]):
            continue
        elif pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        else:
            df_copy[col] = df_copy[col].fillna('').astype(str).str.strip()
            df_copy[col] = df_copy[col].replace('nan', '')
    return df_copy

class OutputFormatter:
    @staticmethod
    def write(result_df: pd.DataFrame, output_dir: str, base_name: str,
              output_format: str) -> (str, str):
        if output_format == 'auto':
            est_size = result_df.shape[0] * result_df.shape[1] * 10
            output_format = 'csv' if est_size > 100 * 1024 * 1024 else 'xlsx'

        formatted_df = safe_format_dataframe(result_df)

        if output_format == 'csv':
            out_path = os.path.join(output_dir, f"{base_name}.csv")
            formatted_df.to_csv(out_path, index=False, encoding='utf-8-sig')
        else:
            out_path = os.path.join(output_dir, f"{base_name}.xlsx")
            write_excel_to_path(formatted_df, out_path, '数据')
        return out_path, output_format
