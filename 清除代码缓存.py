import os
import shutil
from pathlib import Path

def clear_cache():
    """清除所有 __pycache__ 文件夹、.pyc/.pyo 文件，以及临时 Excel 文件，并清空 uploads 和 task_status 目录"""
    project_root = Path(__file__).parent  # 脚本所在目录，即项目根目录
    print(f"正在清理项目缓存: {project_root}")

    # 删除所有 __pycache__ 目录
    for pycache_dir in project_root.rglob("__pycache__"):
        if pycache_dir.is_dir():
            print(f"删除目录: {pycache_dir}")
            shutil.rmtree(pycache_dir, ignore_errors=True)

    # 删除所有 .pyc 和 .pyo 文件
    for ext in ["*.pyc", "*.pyo"]:
        for file in project_root.rglob(ext):
            if file.is_file():
                print(f"删除文件: {file}")
                file.unlink()

    # 清理临时 Excel 文件（可选）
    temp_excel_files = list(project_root.glob("temp*.xlsx"))
    for temp_file in temp_excel_files:
        print(f"删除临时文件: {temp_file}")
        temp_file.unlink()

    # 清空 uploads 目录下的文件（保留子目录结构）
    uploads_dir = project_root / "uploads"
    if uploads_dir.exists() and uploads_dir.is_dir():
        for item in uploads_dir.iterdir():
            if item.is_file():
                print(f"删除上传文件: {item}")
                item.unlink()
            elif item.is_dir():
                # 递归清空子目录中的文件
                for sub_file in item.rglob("*"):
                    if sub_file.is_file():
                        print(f"删除上传文件: {sub_file}")
                        sub_file.unlink()

    # 清空 task_status 目录下的文件（保留子目录结构）
    task_status_dir = project_root / "task_status"
    if task_status_dir.exists() and task_status_dir.is_dir():
        for item in task_status_dir.iterdir():
            if item.is_file():
                print(f"删除任务状态文件: {item}")
                item.unlink()
            elif item.is_dir():
                for sub_file in item.rglob("*"):
                    if sub_file.is_file():
                        print(f"删除任务状态文件: {sub_file}")
                        sub_file.unlink()

    print("清理完成！")

if __name__ == "__main__":
    clear_cache()