# core/excel_writer.py
"""统一 Excel 输出工具：写入 BytesIO 并返回"""
import io
import pandas as pd
from flask import send_file

def write_excel_response(df: pd.DataFrame, filename: str, sheet_name: str = '数据'):
    """将 DataFrame 写入 BytesIO，返回 Flask send_file 响应"""
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl', sheet_name=sheet_name)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def write_excel_to_path(df: pd.DataFrame, path: str, sheet_name: str = '数据'):
    """写入磁盘文件（用于异步任务或后续打包）"""
    df.to_excel(path, index=False, engine='openpyxl', sheet_name=sheet_name)
