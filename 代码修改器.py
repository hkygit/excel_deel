#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
补全测试：为 transposer 添加测试，修复 test_condition_filter.py 的废弃导入。
"""

import os
import shutil

# 新测试文件：tests/test_transposer.py
TRANS_TEST_PATH = os.path.join("tests", "test_transposer.py")
TRANS_TEST_CONTENT = '''# tests/test_transposer.py
import pytest
import pandas as pd
from services.transposer.service import pivot_table


class TestPivotTable:
    def test_pivot_basic_with_index(self):
        df = pd.DataFrame({
            "病种": ["A", "A", "B"],
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300],
            "数量": [1, 2, 3]
        })
        result = pivot_table(
            df,
            index_cols=["病种"],
            category_col="医疗类别",
            value_cols=["金额", "数量"]
        )
        expected_columns = ["病种", "住院金额", "门诊金额", "住院数量", "门诊数量"]
        assert list(result.columns) == expected_columns
        assert result.loc[0, "病种"] == "A"
        assert result.loc[0, "住院金额"] == 100
        assert result.loc[0, "门诊金额"] == 200
        assert result.loc[0, "住院数量"] == 1
        assert result.loc[0, "门诊数量"] == 2
        assert result.loc[1, "病种"] == "B"
        assert result.loc[1, "住院金额"] == 300
        assert pd.isna(result.loc[1, "门诊金额"])
        assert result.loc[1, "住院数量"] == 3
        assert pd.isna(result.loc[1, "门诊数量"])

    def test_pivot_no_index(self):
        df = pd.DataFrame({
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300]
        })
        result = pivot_table(
            df,
            index_cols=None,
            category_col="医疗类别",
            value_cols=["金额"]
        )
        assert "住院金额" in result.columns
        assert "门诊金额" in result.columns
        assert len(result) == 2

    def test_pivot_multiple_value_cols(self):
        df = pd.DataFrame({
            "ID": [1, 1, 2],
            "类别": ["X", "Y", "X"],
            "值1": [10, 20, 30],
            "值2": [100, 200, 300]
        })
        result = pivot_table(
            df,
            index_cols=["ID"],
            category_col="类别",
            value_cols=["值1", "值2"]
        )
        assert "X值1" in result.columns
        assert "Y值1" in result.columns
        assert "X值2" in result.columns
        assert "Y值2" in result.columns
        assert result.loc[0, "ID"] == 1
        assert result.loc[0, "X值1"] == 10
        assert result.loc[0, "Y值1"] == 20
        assert result.loc[0, "X值2"] == 100
        assert result.loc[0, "Y值2"] == 200


class TestTransposerBlueprintIntegration:
    def test_transposer_endpoint(self, client):
        import io
        df = pd.DataFrame({
            "病种": ["A", "A", "B"],
            "医疗类别": ["住院", "门诊", "住院"],
            "金额": [100, 200, 300]
        })
        excel_file = io.BytesIO()
        df.to_excel(excel_file, index=False)
        excel_file.seek(0)

        data = {
            "file": (excel_file, "test.xlsx"),
            "index_cols": "病种",
            "category_col": "医疗类别",
            "value_cols": "金额"
        }
        response = client.post("/transposer/convert", data=data, content_type="multipart/form-data")
        assert response.status_code == 200
        assert "转置结果.xlsx" in response.headers.get("Content-Disposition", "")
'''

# 修复 test_condition_filter.py
CONDITION_FILTER_TEST_PATH = os.path.join("tests", "test_condition_filter.py")

def fix_condition_filter_test():
    if not os.path.exists(CONDITION_FILTER_TEST_PATH):
        print(f"文件不存在，跳过: {CONDITION_FILTER_TEST_PATH}")
        return
    with open(CONDITION_FILTER_TEST_PATH, 'r', encoding='utf-8') as f:
        content = f.read()
    new_content = content.replace("from utils.condition_filter import ConditionFilter, Condition",
                                  "from core.filter_engine import Condition, ConditionFilter")
    new_content = new_content.replace("from utils.condition_filter import Condition",
                                      "from core.filter_engine import Condition")
    if new_content != content:
        backup = CONDITION_FILTER_TEST_PATH + ".bak"
        shutil.copy2(CONDITION_FILTER_TEST_PATH, backup)
        print(f"已备份: {CONDITION_FILTER_TEST_PATH} -> {backup}")
        with open(CONDITION_FILTER_TEST_PATH, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"已更新: {CONDITION_FILTER_TEST_PATH}")
    else:
        print(f"无需修改: {CONDITION_FILTER_TEST_PATH}")

def main():
    print("开始补全测试...")
    # 1. 创建 test_transposer.py
    with open(TRANS_TEST_PATH, 'w', encoding='utf-8') as f:
        f.write(TRANS_TEST_CONTENT)
    print(f"已创建: {TRANS_TEST_PATH}")

    # 2. 修复 test_condition_filter.py
    fix_condition_filter_test()

    print("\n测试补全完成。请手动运行 pytest 验证所有测试通过。")
    print("注意：transposer 的集成测试需要确保 app 和 client fixture 可用（已在 conftest.py 中定义）。")

if __name__ == "__main__":
    main()