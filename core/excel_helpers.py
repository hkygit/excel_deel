"""
Excel 文件辅助函数
"""
import os
import logging
import tempfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

logger = logging.getLogger(__name__)


def normalize_excel_file(filepath: str) -> str:
    """
    将 Excel 文件标准化为 .xlsx 格式
    安全处理文件锁定和XML格式问题
    """
    # 检查文件是否存在
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    # 检查文件扩展名
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xlsx':
        # 已经是 .xlsx 格式，检查文件是否有效
        try:
            # 尝试快速打开检查文件是否有效
            with open(filepath, 'rb') as f:
                header = f.read(4)
                if header != b'PK\x03\x04':
                    logger.warning(f"文件 {filepath} 不是有效的ZIP文件，可能是损坏的Excel文件")
        except Exception as e:
            logger.warning(f"检查文件失败: {e}")

    temp_dir = tempfile.gettempdir()
    base_name = os.path.basename(filepath)
    output_name = f"normalized_{os.path.splitext(base_name)[0]}_{os.urandom(4).hex()}.xlsx"
    output_path = os.path.join(temp_dir, output_name)

    try:
        # 方法1: 首先尝试用openpyxl读取
        try:
            logger.info(f"尝试用openpyxl读取文件: {filepath}")
            # 使用只读模式减少内存使用
            wb = load_workbook(
                filename=filepath,
                data_only=True,
                read_only=True,  # 只读模式，减少内存使用
                keep_vba=False
            )

            # 创建新的工作簿用于写入
            from openpyxl import Workbook
            wb_new = Workbook()
            wb_new.remove(wb_new.active)  # 删除默认的工作表

            # 复制所有工作表
            for ws in wb.worksheets:
                ws_new = wb_new.create_sheet(title=ws.title)
                # 只复制有限的行，避免内存问题
                max_row = min(ws.max_row, 100000)  # 限制行数
                max_col = min(ws.max_column, 100)  # 限制列数

                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        ws_new.cell(row=row, column=col, value=cell.value)

            wb.close()  # 关闭只读工作簿
            wb_new.save(output_path)
            wb_new.close()

            logger.info(f"已用openpyxl转换文件: {output_path}")
            return output_path

        except Exception as e:
            logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")

            # 方法2: 使用pandas读取
            try:
                # 根据文件扩展名选择engine
                engine = None
                if ext in ['.xls', '.xlsm']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

                # 读取Excel文件
                df = pd.read_excel(
                    filepath,
                    dtype=str,  # 全部作为字符串读取
                    na_filter=False,  # 不过滤空值
                    engine=engine,
                    header=None  # 不指定标题行
                )

                # 清理数据
                df = df.fillna('')

                # 保存为新的Excel文件
                df.to_excel(
                    output_path,
                    index=False,
                    header=False,
                    engine='openpyxl'
                )

                logger.info(f"已用pandas转换文件: {output_path}")
                return output_path

            except Exception as pd_e:
                logger.error(f"pandas 读取也失败: {pd_e}")

                # 方法3: 尝试修复文件
                try:
                    # 创建一个新的空Excel文件
                    from openpyxl import Workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Data"
                    ws['A1'] = "错误: 无法读取原始文件"
                    ws['A2'] = f"原始错误: {str(pd_e)[:100]}"
                    wb.save(output_path)
                    logger.warning(f"创建了空的修复文件: {output_path}")
                    return output_path

                except Exception as create_e:
                    logger.error(f"创建修复文件也失败: {create_e}")
                    raise ValueError(f"无法读取或转换 Excel 文件。请确保文件是有效的 Excel 文件。原始错误: {pd_e}")

    except Exception as final_e:
        logger.error(f"所有转换方法都失败: {final_e}")
        # 如果输出路径已创建，尝试清理
        if os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except:
                pass
        raise ValueError(f"Excel 文件处理失败: {final_e}")


def safe_cleanup_dataframe(df: pd.DataFrame, auto_convert_numeric: bool = True) -> pd.DataFrame:
    """
    安全清理 DataFrame
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    for col in df.columns:
        # 清理字符串
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            # 将空字符串转为空值
            df[col] = df[col].replace('', pd.NA)
            df[col] = df[col].replace('nan', pd.NA)
            df[col] = df[col].replace('None', pd.NA)
            df[col] = df[col].replace('null', pd.NA)

        # 自动转换数值
        if auto_convert_numeric and df[col].dtype == 'object':
            try:
                numeric_series = pd.to_numeric(df[col], errors='coerce')
                if numeric_series.notna().any():
                    df[col] = numeric_series
            except Exception:
                pass

    return df


def build_column_name_mapping(ws, header_row: int):
    """
    构建列名到列索引的映射
    """
    mapping = {}
    max_col = min(ws.max_column, 100)  # 限制列数

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        if cell and cell.value:
            col_name = str(cell.value).strip()
            if col_name:
                mapping[col_name] = col_idx

    return mapping


def safe_get_cell_value(ws, row: int, col: int):
    """
    安全获取单元格值
    """
    try:
        if row > ws.max_row or col > ws.max_column:
            return None
        cell = ws.cell(row=row, column=col)
        return cell.value
    except Exception:
        return None


def safe_write_cell(ws, row: int, col: int, value):
    """
    安全写入单元格值
    """
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
    except Exception as e:
        logger.warning(f"写入单元格失败 ({row}, {col}): {e}")


def clear_merges_in_column(ws, col_idx: int):
    """
    清除指定列的合并单元格
    """
    try:
        ranges_to_remove = []
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col <= col_idx <= merged_range.max_col:
                ranges_to_remove.append(str(merged_range))

        for range_str in ranges_to_remove:
            ws.unmerge_cells(range_str)
    except Exception as e:
        logger.warning(f"清除合并单元格失败: {e}")


def safe_delete_file(filepath: str, max_retries: int = 3):
    """
    安全删除文件，处理文件被占用的情况
    """
    if not filepath or not os.path.exists(filepath):
        return True

    for i in range(max_retries):
        try:
            os.unlink(filepath)
            logger.info(f"成功删除文件: {filepath}")
            return True
        except PermissionError as e:
            if i < max_retries - 1:
                import time
                time.sleep(0.1 * (i + 1))  # 指数退避
                continue
            logger.warning(f"删除文件失败(被占用): {filepath}, 错误: {e}")
            return False
        except Exception as e:
            logger.warning(f"删除文件失败: {filepath}, 错误: {e}")
            return False

    return False
