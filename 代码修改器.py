#!/usr/bin/env python3
"""
一键生成各模块的 POML 配置文件（.txt），存放于 POML/ 目录下。
每个配置文件包含该模块运行所需的所有相关文件（核心、蓝图、服务、测试、前端等）。
"""

import os
from pathlib import Path

# ========== 定义各模块的配置文件内容 ==========
CONFIGS = {
    "calc.txt": """# calc模块相关文件
blueprints/calculator.py
services/excel/calculator_processor.py
services/excel/compute_engine.py
services/excel/aggregator.py
services/data/reader.py
services/data/writer.py
services/data/formatter.py
services/data/file_service.py
core/expression_engine.py
core/excel_utils.py
core/data_access_layer.py
core/header_detection.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/calculator/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_calculator.py
tests/test_integration_calculator.py
""",

    "compare.txt": """# compare模块相关文件
blueprints/comparator.py
services/excel/comparator.py
services/data/file_service.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
schemas/comparator.py
templates/comparator/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_comparator.py
tests/test_integration_comparator.py
""",

    "match.txt": """# match模块相关文件
blueprints/matcher.py
services/excel/matcher.py
services/data/reader.py
core/excel_utils.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/matcher/index.html
templates/debug.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_integration_matcher.py
""",

    "summary.txt": """# summary模块相关文件
blueprints/summarizer.py
services/excel/compute_engine.py
services/excel/aggregator.py
services/data/reader.py
services/data/writer.py
services/data/formatter.py
services/data/file_service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/header_detection.py
core/exceptions.py
core/config.py
core/di.py
templates/summarizer/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_integration_summarizer.py
""",

    "splitter.txt": """# splitter模块相关文件
blueprints/splitter.py
core/filter_engine.py
core/validators.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/splitter/index.html
templates/base.html
static/css/style.css
static/js/main.js
tests/test_splitter.py
tests/test_integration_splitter.py
""",

    "field_parser.txt": """# field_parser模块相关文件
blueprints/field_parser.py
services/field_parser/service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/field_parser/index.html
templates/base.html
static/css/style.css
static/js/field_parser.js
static/js/main.js
tests/test_field_parser.py
tests/test_integration_field_parser.py
""",

    "checker.txt": """# checker模块相关文件
blueprints/checker.py
services/checker/__init__.py
services/checker/service.py
services/checker/strategies.py
services/checker/constants.py
services/data/stream_reader.py
services/data/row_counter.py
services/data/report_writer.py
core/filter_engine.py
core/progress_task.py
core/sqlite_buffer.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/checker/index.html
templates/base.html
static/css/checker.css
static/css/style.css
static/js/checker.js
static/js/main.js
tests/test_checker_service.py
tests/test_integration_checker.py
""",

    "convert.txt": """# convert模块相关文件
blueprints/convert.py
services/convert/__init__.py
services/convert/service.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/convert/index.html
templates/base.html
static/css/style.css
static/js/convert.js
static/js/main.js
tests/test_convert_service.py
tests/test_integration_convert.py
""",

    "team_report.txt": """# team_report模块相关文件
blueprints/team_report.py
services/team_report/__init__.py
services/team_report/service.py
services/team_report/writer.py
services/common/excel_styles.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/team_report/index.html
templates/base.html
static/css/style.css
static/js/team_report.js
static/js/main.js
tests/test_team_report_service.py
tests/test_integration_team_report.py
tests/test_team_report_long_running.py
""",

    "transposer.txt": """# transposer模块相关文件
blueprints/transposer.py
services/transposer/service.py
core/response.py
core/decorators.py
core/file_naming.py
core/csv_utils.py
core/exceptions.py
core/config.py
core/di.py
templates/transposer/index.html
templates/base.html
static/css/style.css
static/js/transposer.js
static/js/main.js
tests/test_transposer.py
""",

    "subtotal.txt": """# subtotal模块相关文件
blueprints/subtotal.py
services/subtotal/__init__.py
services/subtotal/engine.py
core/header_detection.py
core/csv_utils.py
core/task/__init__.py
core/task/base.py
core/task/decorator.py
core/task/manager.py
core/task/state.py
core/task/storage.py
core/response.py
core/decorators.py
core/file_naming.py
core/exceptions.py
core/config.py
core/di.py
templates/subtotal/index.html
templates/base.html
static/css/style.css
static/js/subtotal.js
static/js/main.js
tests/test_subtotal_engine.py
tests/test_integration_subtotal.py
""",
}

def main():
    # 获取项目根目录（脚本所在目录）
    base_dir = Path(__file__).resolve().parent
    poml_dir = base_dir / "POML"
    poml_dir.mkdir(exist_ok=True)
    print(f"✓ POML 目录已创建/确认: {poml_dir}")

    for filename, content in CONFIGS.items():
        filepath = poml_dir / filename
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"✓ 已生成: {filepath}")

    print("\n✅ 所有模块配置文件已生成完毕。")
    print("现在可以运行 python 整合git喂给ai.py 并选择选项 2，输入对应的 .txt 文件名来导出模块代码快照。")

if __name__ == "__main__":
    main()