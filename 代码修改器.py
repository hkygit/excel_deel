#!/usr/bin/env python3
"""
静态检查：验证所有生成 Excel 的蓝图是否使用了统一的文件命名函数
"""
import os
import ast
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
BLUEPRINTS_DIR = os.path.join(ROOT, 'blueprints')

# 需要检查的蓝图模块及其功能描述（文件名 -> 描述）
TARGET_MODULES = {
    'calculator.py': '列计算/聚合回填',
    'comparator.py': '文件对比',
    'matcher.py': '匹配填充',
    'summarizer.py': '汇总工具',
    'splitter.py': '文件拆分',
    'subtotal.py': '数据小计',
    'field_parser.py': '字段映射导出',
    'checker.py': '文件校验下载',
    'convert.py': '格式转换下载',
    'team_report.py': '车队日报下载',
    'transposer.py': '行列转置',
}

def check_import(node):
    """检查 AST 中是否导入了 generate_output_filename"""
    for n in ast.walk(node):
        if isinstance(n, ast.ImportFrom):
            if n.module == 'core.file_naming':
                for alias in n.names:
                    if alias.name == 'generate_output_filename':
                        return True
    return False

def check_usage_in_function(node, func_name):
    """检查指定函数内是否调用了 generate_output_filename"""
    for n in ast.walk(node):
        if isinstance(n, ast.FunctionDef) and n.name == func_name:
            # 遍历函数体中的调用
            for sub in ast.walk(n):
                if isinstance(sub, ast.Call):
                    # 直接调用 generate_output_filename(...)
                    if isinstance(sub.func, ast.Name) and sub.func.id == 'generate_output_filename':
                        return True, '直接调用'
                    # 从模块调用，如 core.file_naming.generate_output_filename
                    if isinstance(sub.func, ast.Attribute):
                        if sub.func.attr == 'generate_output_filename':
                            # 检查是否赋值给了 download_name
                            return True, '模块调用'
            # 如果没有找到，可能是在 send_file 的 download_name 中使用
            for sub in ast.walk(n):
                if isinstance(sub, ast.keyword) and sub.arg == 'download_name':
                    if isinstance(sub.value, ast.Call) and isinstance(sub.value.func, ast.Name) and sub.value.func.id == 'generate_output_filename':
                        return True, 'download_name参数'
                    if isinstance(sub.value, ast.Call) and isinstance(sub.value.func, ast.Attribute) and sub.value.func.attr == 'generate_output_filename':
                        return True, 'download_name参数'
            return False, ''
    return False, ''

def check_module(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    try:
        tree = ast.parse(content)
    except SyntaxError as e:
        return {'imported': False, 'used': False, 'error': str(e)}

    imported = check_import(tree)

    # 查找可能的下载函数：通常为 download、export、或者直接返回 send_file 的视图函数
    used = False
    detail = ''
    # 检查所有函数
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef):
            # 常见的下载/导出函数名
            if node.name in ('download', 'export', 'convert', 'process', 'generate', 'debug'):
                # 在这些函数中查找是否有 generate_output_filename 的使用
                # 遍历函数体中的 send_file 调用
                for sub in ast.walk(node):
                    if isinstance(sub, ast.Call) and isinstance(sub.func, ast.Name) and sub.func.id == 'send_file':
                        # 检查 download_name 参数
                        for kw in sub.keywords:
                            if kw.arg == 'download_name':
                                if isinstance(kw.value, ast.Call):
                                    if isinstance(kw.value.func, ast.Name) and kw.value.func.id == 'generate_output_filename':
                                        used = True
                                        detail = f'在函数 {node.name} 的 download_name 中使用'
                                break
                if not used:
                    # 也可能是先调用 generate_output_filename 再作为变量传给 send_file
                    for sub in ast.walk(node):
                        if isinstance(sub, ast.Assign):
                            for target in sub.targets:
                                if isinstance(target, ast.Name):
                                    if isinstance(sub.value, ast.Call) and isinstance(sub.value.func, ast.Name) and sub.value.func.id == 'generate_output_filename':
                                        used = True
                                        detail = f'在函数 {node.name} 中赋值给变量 {target.id}'
    return {'imported': imported, 'used': used, 'detail': detail}

def main():
    print("=" * 60)
    print("统一文件命名静态检查")
    print("=" * 60)
    results = {}
    for filename, desc in TARGET_MODULES.items():
        filepath = os.path.join(BLUEPRINTS_DIR, filename)
        if not os.path.exists(filepath):
            results[filename] = {'desc': desc, 'status': '文件不存在'}
            continue
        check_result = check_module(filepath)
        results[filename] = {**check_result, 'desc': desc}

    # 输出表格
    print(f"{'模块':<20} {'功能':<16} {'导入':<6} {'使用':<6} {'详情'}")
    print("-" * 70)
    for fname, info in results.items():
        if 'status' in info:
            print(f"{fname:<20} {info['desc']:<16} {'❌':<6} {'❌':<6} {info['status']}")
        else:
            imp = '✅' if info['imported'] else '❌'
            use = '✅' if info['used'] else '❌'
            detail = info.get('detail', '')
            print(f"{fname:<20} {info['desc']:<16} {imp:<6} {use:<6} {detail}")

    print("\n总结：")
    all_ok = all(
        'status' not in info and info.get('imported') and info.get('used')
        for info in results.values()
    )
    if all_ok:
        print("✅ 所有模块均已导入并使用统一命名函数。")
    else:
        print("⚠️ 以下模块需要修复：")
        for fname, info in results.items():
            if 'status' in info:
                print(f"  - {fname}: {info['status']}")
            elif not info.get('imported'):
                print(f"  - {fname}: 未导入 generate_output_filename")
            elif not info.get('used'):
                print(f"  - {fname}: 未在下载/导出中使用 generate_output_filename")

if __name__ == '__main__':
    main()