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


# Excel 单表硬限制（参考 Office Open XML 规范）
MAX_EXCEL_ROWS = 1_048_576   # 2^20
MAX_EXCEL_COLS = 16_384      # 2^14


class ExcelNormalizationError(ValueError):
    """
    Excel 文件规范化过程中无法安全处理时抛出。

    当前触发场景：
    - 单表行数 / 列数超过设定的上限（不再静默截断）

    继承自 ValueError 以保持向后兼容：调用方若只 except ValueError 仍能捕获。
    """


def _check_size_limits(rows: int, cols: int, sheet_label: str,
                       max_rows: int, max_cols: int) -> None:
    """
    检查工作表尺寸是否超过上限，超过则抛 ExcelNormalizationError。

    这里刻意采用「抛异常」策略，而不是「静默截断」——
    静默截断会造成用户数据丢失而不被察觉，属于数据完整性事故。
    """
    if rows and rows > max_rows:
        raise ExcelNormalizationError(
            f"{sheet_label} 行数 {rows:,} 超过上限 {max_rows:,}；"
            f"为避免静默截断数据，处理已中止。"
            f"建议先拆分/抽样文件，或通过 max_rows 参数显式提高上限。"
        )
    if cols and cols > max_cols:
        raise ExcelNormalizationError(
            f"{sheet_label} 列数 {cols:,} 超过上限 {max_cols:,}；"
            f"为避免静默截断数据，处理已中止。"
            f"建议先拆分/抽样文件，或通过 max_cols 参数显式提高上限。"
        )


def normalize_excel_file(
    filepath: str,
    max_rows: int = MAX_EXCEL_ROWS,
    max_cols: int = MAX_EXCEL_COLS,
) -> str:
    """
    将 Excel 文件标准化为 .xlsx 格式。

    与旧版相比的重要变化：
    - 移除了原先的静默截断（100000 行 / 100 列上限）
    - 当文件超出 max_rows / max_cols 时，抛出 ExcelNormalizationError
    - 正常路径下完整复制所有行、列，不再丢数据

    Args:
        filepath: 原始 Excel 文件路径。
        max_rows: 单表允许的最大行数（默认 Excel 硬限制 1,048,576）。
        max_cols: 单表允许的最大列数（默认 Excel 硬限制 16,384）。

    Returns:
        规范化后的临时 .xlsx 文件路径。

    Raises:
        FileNotFoundError: 源文件不存在。
        ExcelNormalizationError: 超过行/列上限。
        ValueError: 所有读取方式都失败。
    """
    # 检查文件是否存在
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xlsx':
        try:
            with open(filepath, 'rb') as f:
                header = f.read(4)
                if header != b'PK\x03\x04':
                    logger.warning(
                        f"文件 {filepath} 不是有效的ZIP文件，可能是损坏的Excel文件"
                    )
        except Exception as e:
            logger.warning(f"检查文件失败: {e}")

    temp_dir = tempfile.gettempdir()
    base_name = os.path.basename(filepath)
    output_name = (
        f"normalized_{os.path.splitext(base_name)[0]}_{os.urandom(4).hex()}.xlsx"
    )
    output_path = os.path.join(temp_dir, output_name)

    def _safe_unlink(path: str) -> None:
        if path and os.path.exists(path):
            try:
                os.unlink(path)
            except Exception as cleanup_err:
                logger.debug(f"清理临时文件失败: {path} - {cleanup_err}")

    try:
        # 方法1: 首先尝试用 openpyxl 读取
        try:
            logger.info(f"尝试用openpyxl读取文件: {filepath}")
            wb = load_workbook(
                filename=filepath,
                data_only=True,
                read_only=True,      # 只读模式，减小内存占用
                keep_vba=False,
            )

            # === 关键改动 1：先检查上限，再决定是否处理 ===
            # 在 read_only 模式下，max_row / max_column 来自文件元数据，
            # 不需要真正加载全部单元格即可获得，可提前拦截超大文件。
            for ws in wb.worksheets:
                rows = ws.max_row or 0
                cols = ws.max_column or 0
                if rows > max_rows or cols > max_cols:
                    wb.close()
                    _safe_unlink(output_path)
                    # 抛出异常（不会进入 pandas 回退）
                    _check_size_limits(
                        rows, cols,
                        sheet_label=f"工作表 '{ws.title}'",
                        max_rows=max_rows, max_cols=max_cols,
                    )

            # 创建新的工作簿用于写入
            from openpyxl import Workbook
            wb_new = Workbook()
            wb_new.remove(wb_new.active)  # 删除默认的工作表

            # === 关键改动 2：完整复制所有行列，不再截断 ===
            for ws in wb.worksheets:
                ws_new = wb_new.create_sheet(title=ws.title)
                # iter_rows(values_only=True) 是 read_only 模式下的推荐迭代方式，
                # 相比 cell-by-cell 访问更快、内存更省。
                for row in ws.iter_rows(values_only=True):
                    ws_new.append(list(row))

            wb.close()
            wb_new.save(output_path)
            wb_new.close()

            logger.info(f"已用openpyxl转换文件: {output_path}")
            return output_path

        except ExcelNormalizationError:
            # 尺寸超限属于「不可恢复」错误，直接向上传播，禁止回退到 pandas
            # （pandas 会把整个文件读入内存，反而更危险）
            raise
        except Exception as e:
            logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")

            # 方法2: 使用 pandas 读取（对损坏文件 / 老格式的兜底）
            try:
                engine = None
                if ext in ['.xls', '.xlsm']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

                df = pd.read_excel(
                    filepath,
                    dtype=str,          # 全部作为字符串读取
                    na_filter=False,    # 不过滤空值
                    engine=engine,
                    header=None,        # 不指定标题行
                )

                # === 关键改动 3：pandas 分支同样检查上限 ===
                n_rows = df.shape[0]
                n_cols = df.shape[1]
                if n_rows > max_rows or n_cols > max_cols:
                    _safe_unlink(output_path)
                    _check_size_limits(
                        n_rows, n_cols,
                        sheet_label=f"文件 '{os.path.basename(filepath)}'",
                        max_rows=max_rows, max_cols=max_cols,
                    )

                df = df.fillna('')
                df.to_excel(
                    output_path,
                    index=False,
                    header=False,
                    engine='openpyxl',
                )

                logger.info(f"已用pandas转换文件: {output_path}")
                return output_path

            except ExcelNormalizationError:
                raise
            except Exception as pd_e:
                logger.error(f"pandas 读取也失败: {pd_e}")

                # 方法3: 生成一个说明性的空文件（保留原有兜底行为）
                try:
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
                    raise ValueError(
                        f"无法读取或转换 Excel 文件。请确保文件是有效的 Excel 文件。"
                        f"原始错误: {pd_e}"
                    )

    except Exception as final_e:
        logger.error(f"所有转换方法都失败: {final_e}")
        _safe_unlink(output_path)
        raise


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
