"""
Excel 文件辅助函数。

P0 变更：
- normalize_excel_file 对已经是有效 .xlsx 的文件直接返回原路径，
  避免无意义的"复制到临时文件再读回"。
- 新增 _validate_xlsx 做轻量校验（文件头 + ZIP 结构 + openpyxl 尺寸检查）。
"""
import os
import logging
import tempfile
import zipfile
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
    """尺寸超限时抛 ExcelNormalizationError。"""
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


def _validate_xlsx(filepath: str, max_rows: int, max_cols: int) -> bool:
    """快速校验 .xlsx 是否有效且尺寸合规。

    返回:
        True  - 文件有效且尺寸合规（可安全地返回原路径）
        False - 文件无效（需走转换路径）
    抛出:
        ExcelNormalizationError - 尺寸超限（不进入转换路径）

    不创建任何临时文件，不修改原文件。
    """
    # 1) 文件头检查（ZIP magic）
    try:
        with open(filepath, 'rb') as f:
            if f.read(4) != b'PK\x03\x04':
                return False
    except OSError:
        return False

    # 2) ZIP 结构检查（比 openpyxl 更轻量，不解析样式 / 关系）
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            names = zf.namelist()
            # 有效 xlsx 至少包含 xl/ 目录下的内容
            if not any(n.startswith('xl/') for n in names):
                return False
    except (zipfile.BadZipFile, OSError):
        return False

    # 3) openpyxl 打开 + 尺寸检查（read_only 模式下 max_row / max_column
    #    从 XML 元数据读取，不会真的全量扫描单元格）
    try:
        wb = load_workbook(
            filename=filepath,
            data_only=True,
            read_only=True,
            keep_vba=False,
        )
    except Exception:
        return False

    try:
        for ws in wb.worksheets:
            rows = ws.max_row or 0
            cols = ws.max_column or 0
            if rows > max_rows or cols > max_cols:
                _check_size_limits(
                    rows, cols,
                    sheet_label=f"工作表 '{ws.title}'",
                    max_rows=max_rows, max_cols=max_cols,
                )
    finally:
        wb.close()

    return True


def normalize_excel_file(
    filepath: str,
    max_rows: int = MAX_EXCEL_ROWS,
    max_cols: int = MAX_EXCEL_COLS,
) -> str:
    """确保文件是可被下游正常读写的 .xlsx。

    P0 语义变化：
    - 对**有效的 .xlsx**：直接返回原路径（零复制），调用方需保证只读不改
    - 对 .xls / .xlsm / 损坏 .xlsx：走转换路径，返回新的临时文件
    - 尺寸超限：抛 ExcelNormalizationError（不静默截断）

    Args:
        filepath: 原始文件路径
        max_rows: 单表允许的最大行数（默认 Excel 硬限制）
        max_cols: 单表允许的最大列数（默认 Excel 硬限制）

    Returns:
        可安全读写的 .xlsx 路径（可能是原路径，也可能是新的临时文件）

    Raises:
        FileNotFoundError: 源文件不存在
        ExcelNormalizationError: 超过行/列上限
        ValueError: 所有读取方式都失败
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()

    # ===== P0：有效 .xlsx 直接返回原路径，零复制 =====
    if ext == '.xlsx':
        try:
            if _validate_xlsx(filepath, max_rows, max_cols):
                logger.info(
                    f"文件已是有效 .xlsx，直接使用原路径: {filepath}"
                )
                return filepath
        except ExcelNormalizationError:
            # 尺寸超限：立即向上抛，**不走转换路径**
            # 原因：pandas / openpyxl 的全量读更容易 OOM，及早拦截更安全
            raise
        except Exception as e:  # noqa: BLE001
            logger.warning(
                f".xlsx 有效性检测异常，降级为转换路径: {e}"
            )

    # ===== 以下为需要转换的场景（.xls / .xlsm / 损坏 .xlsx）=====
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
            except Exception as cleanup_err:  # noqa: BLE001
                logger.debug(f"清理临时文件失败: {path} - {cleanup_err}")

    try:
        # 方法1: 用 openpyxl 读取并另存
        try:
            logger.info(f"尝试用openpyxl读取文件: {filepath}")
            wb = load_workbook(
                filename=filepath,
                data_only=True,
                read_only=True,
                keep_vba=False,
            )

            # 尺寸检查（覆盖 .xls / .xlsm 等 _validate_xlsx 未校验的场景）
            for ws in wb.worksheets:
                rows = ws.max_row or 0
                cols = ws.max_column or 0
                if rows > max_rows or cols > max_cols:
                    wb.close()
                    _safe_unlink(output_path)
                    _check_size_limits(
                        rows, cols,
                        sheet_label=f"工作表 '{ws.title}'",
                        max_rows=max_rows, max_cols=max_cols,
                    )

            # 完整复制所有行列（read_only 模式下用 iter_rows 最快）
            from openpyxl import Workbook
            wb_new = Workbook()
            wb_new.remove(wb_new.active)

            for ws in wb.worksheets:
                ws_new = wb_new.create_sheet(title=ws.title)
                for row in ws.iter_rows(values_only=True):
                    ws_new.append(list(row))

            wb.close()
            wb_new.save(output_path)
            wb_new.close()

            logger.info(f"已用openpyxl转换文件: {output_path}")
            return output_path

        except ExcelNormalizationError:
            raise
        except Exception as e:  # noqa: BLE001
            logger.warning(f"openpyxl 转换失败: {e}，回退到 pandas 方式")

            # 方法2: pandas 兜底
            try:
                engine = None
                if ext in ['.xls', '.xlsm']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

                df = pd.read_excel(
                    filepath,
                    dtype=str,
                    na_filter=False,
                    engine=engine,
                    header=None,
                )

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
                    output_path, index=False, header=False,
                    engine='openpyxl',
                )

                logger.info(f"已用pandas转换文件: {output_path}")
                return output_path

            except ExcelNormalizationError:
                raise
            except Exception as pd_e:  # noqa: BLE001
                logger.error(f"pandas 读取也失败: {pd_e}")

                # 方法3: 生成占位文件（保留原有兜底行为）
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
                except Exception as create_e:  # noqa: BLE001
                    logger.error(f"创建修复文件也失败: {create_e}")
                    raise ValueError(
                        f"无法读取或转换 Excel 文件。请确保文件是有效的 Excel 文件。"
                        f"原始错误: {pd_e}"
                    )

    except Exception as final_e:
        logger.error(f"所有转换方法都失败: {final_e}")
        _safe_unlink(output_path)
        raise


def safe_cleanup_dataframe(df: pd.DataFrame,
                           auto_convert_numeric: bool = True) -> pd.DataFrame:
    """安全清理 DataFrame。"""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('', pd.NA)
            df[col] = df[col].replace('nan', pd.NA)
            df[col] = df[col].replace('None', pd.NA)
            df[col] = df[col].replace('null', pd.NA)

        if auto_convert_numeric and df[col].dtype == 'object':
            try:
                numeric_series = pd.to_numeric(df[col], errors='coerce')
                if numeric_series.notna().any():
                    df[col] = numeric_series
            except Exception:
                pass

    return df


def build_column_name_mapping(ws, header_row: int):
    """构建列名到列索引的映射。"""
    mapping = {}
    max_col = min(ws.max_column, 100)

    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        if cell and cell.value:
            col_name = str(cell.value).strip()
            if col_name:
                mapping[col_name] = col_idx

    return mapping


def safe_get_cell_value(ws, row: int, col: int):
    """安全获取单元格值。"""
    try:
        if row > ws.max_row or col > ws.max_column:
            return None
        cell = ws.cell(row=row, column=col)
        return cell.value
    except Exception:
        return None


def safe_write_cell(ws, row: int, col: int, value):
    """安全写入单元格值。"""
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
    except Exception as e:  # noqa: BLE001
        logger.warning(f"写入单元格失败 ({row}, {col}): {e}")


def clear_merges_in_column(ws, col_idx: int):
    """清除指定列的合并单元格。"""
    try:
        ranges_to_remove = []
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_col <= col_idx <= merged_range.max_col:
                ranges_to_remove.append(str(merged_range))

        for range_str in ranges_to_remove:
            ws.unmerge_cells(range_str)
    except Exception as e:  # noqa: BLE001
        logger.warning(f"清除合并单元格失败: {e}")


def safe_delete_file(filepath: str, max_retries: int = 3):
    """安全删除文件，处理文件被占用的情况。"""
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
                time.sleep(0.1 * (i + 1))
                continue
            logger.warning(f"删除文件失败(被占用): {filepath}, 错误: {e}")
            return False
        except Exception as e:  # noqa: BLE001
            logger.warning(f"删除文件失败: {filepath}, 错误: {e}")
            return False

    return False
