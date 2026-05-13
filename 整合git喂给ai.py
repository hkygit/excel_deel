#!/usr/bin/env python3
"""
一键导出 Git 仓库中未被 .gitignore 忽略的代码文件为 Markdown（供 AI 分析）
交互版：运行后选择导出范围（整个仓库 / 模块 / 目录），自动生成对应快照。

用法：
    python 整合git喂给ai.py              # 交互模式
    python 整合git喂给ai.py --help       # 查看命令行参数（仍支持非交互运行）
"""

import subprocess
import os
import sys
from pathlib import Path

# ---------------------- 可自定义常量 ----------------------
TEXT_EXTENSIONS = {
    # 后端
    ".py", ".pyw", ".pyx", ".pxd", ".pxi",
    ".java", ".kt", ".kts",
    ".c", ".cpp", ".cc", ".cxx", ".h", ".hpp", ".hxx", ".inl",
    ".cs", ".vb",
    ".go", ".rs", ".swift", ".m", ".mm",
    ".rb", ".rbw",
    ".php", ".phtml", ".php3", ".php4", ".php5", ".phps",
    ".pl", ".pm",
    ".sh", ".bash", ".zsh", ".fish", ".ps1", ".bat", ".cmd",
    ".r", ".R", ".Rmd",
    ".scala", ".clj", ".cljs", ".cljc", ".edn",
    ".lua", ".tcl",
    # 前端 & 标记模板
    ".js", ".mjs", ".cjs", ".ts", ".jsx", ".tsx",
    ".vue", ".svelte", ".astro",
    ".html", ".htm", ".xhtml",
    ".css", ".scss", ".sass", ".less", ".styl",
    ".xml", ".xsl", ".xslt", ".svg",
    # 配置 / 数据 / 文档
    ".json", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".conf", ".properties",
    ".env", ".env.example", ".editorconfig", ".gitignore", ".gitattributes",
    ".md", ".markdown", ".rst", ".txt", ".csv", ".tsv",
    ".sql", ".graphql", ".gql", ".proto", ".avsc",
}
MAX_FILE_BYTES = 500 * 1024  # 500 KB
OUTPUT_FILENAME_BASE = "project_code_for_ai"  # 基础文件名，实际输出加模块后缀
# --------------------------------------------------------

def find_repo_root(start: Path) -> Path:
    """向上查找包含 .git 的目录"""
    current = start.resolve()
    for _ in range(50):
        if (current / ".git").exists():
            return current
        parent = current.parent
        if parent == current:
            break
        current = parent
    print("错误：未找到 Git 仓库根目录（没有 .git 文件夹）。", file=sys.stderr)
    sys.exit(1)

def run_git_ls_files(repo_root: Path) -> list[str]:
    """调用 git ls-files 返回所有未被忽略的文件"""
    try:
        result = subprocess.run(
            ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
            cwd=str(repo_root),
            capture_output=True,
            text=True,
            check=True,
        )
        lines = result.stdout.strip().split("\n")
        return [line.strip() for line in lines if line.strip()]
    except subprocess.CalledProcessError as e:
        print(f"git ls-files 执行失败: {e.stderr}", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError:
        print("错误：系统中未找到 git 命令。请安装 Git 并确保在 PATH 中。", file=sys.stderr)
        sys.exit(1)

def is_text_extension(filename: str) -> bool:
    """通过文件名 / 扩展名判断是否为代码 / 文本文件"""
    name = os.path.basename(filename).lower()
    if name in {
        "makefile", "dockerfile", "jenkinsfile", "vagrantfile",
        "gemfile", "rakefile", "cmakelists.txt"
    }:
        return True
    if "." in name:
        base_part = name.split(".", 1)[0]
        if base_part in {"makefile", "dockerfile"}:
            return True
    ext = os.path.splitext(filename)[1].lower()
    return ext in TEXT_EXTENSIONS

def is_likely_binary(filepath: Path) -> bool:
    """读取文件前 1024 字节，若含 null 字节则为二进制"""
    try:
        with open(filepath, "rb") as f:
            chunk = f.read(1024)
        return b"\x00" in chunk
    except Exception:
        return True

def read_text_safe(filepath: Path, max_bytes: int) -> str | None:
    """安全读取文本文件，超过大小或错误则返回 None"""
    try:
        size = filepath.stat().st_size
        if size > max_bytes:
            return None
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except Exception:
        return None

def get_lang_hint(filename: str) -> str:
    """返回适用于 Markdown 代码块的语言标识"""
    ext = os.path.splitext(filename)[1].lower()
    mapping = {
        ".py": "python", ".pyw": "python",
        ".js": "javascript", ".mjs": "javascript", ".cjs": "javascript",
        ".ts": "typescript", ".jsx": "jsx", ".tsx": "tsx",
        ".java": "java", ".kt": "kotlin", ".kts": "kotlin",
        ".c": "c", ".cpp": "cpp", ".cc": "cpp", ".h": "cpp", ".hpp": "cpp",
        ".cs": "csharp", ".go": "go", ".rs": "rust", ".swift": "swift",
        ".rb": "ruby", ".php": "php", ".sh": "bash", ".bash": "bash",
        ".zsh": "bash", ".ps1": "powershell", ".bat": "batch",
        ".sql": "sql", ".json": "json", ".yaml": "yaml", ".yml": "yaml",
        ".toml": "toml", ".xml": "xml", ".svg": "xml",
        ".md": "markdown", ".css": "css", ".scss": "scss", ".less": "less",
        ".html": "html", ".vue": "html", ".svelte": "html",
    }
    return mapping.get(ext, "")

def sanitize_module_name(module: str) -> str:
    """将模块路径转为安全的文件名片段"""
    safe = module.strip("/\\").replace("/", "_").replace("\\", "_")
    return safe if safe else "full"

def filter_files(files: list[str], module_filter: str | None) -> tuple[list[str], str | None]:
    """根据过滤条件筛选文件列表，返回 (过滤后文件列表, 模块字符串用于命名)"""
    if module_filter is None:
        return files, None

    module_filter = module_filter.strip().replace("\\", "/")
    if not module_filter:
        return files, None

    prefix = module_filter if module_filter.endswith("/") else module_filter + "/"
    filtered = []
    for f in files:
        if f == module_filter or f.startswith(prefix):
            filtered.append(f)
    return filtered, module_filter

def export_to_markdown(repo_root: Path, files: list[str], module_str: str | None):
    """执行导出操作"""
    # 构建输出文件名
    if module_str:
        suffix = sanitize_module_name(module_str)
        output_name = f"{OUTPUT_FILENAME_BASE}_{suffix}.md"
    else:
        output_name = f"{OUTPUT_FILENAME_BASE}.md"
    output_path = repo_root / output_name

    count = 0
    with open(output_path, "w", encoding="utf-8") as md:
        md.write(f"# 项目代码快照\n\n")
        md.write(f"仓库: {repo_root.name}\n")
        md.write(f"过滤条件: {module_str if module_str else '全部'}\n")
        md.write(f"匹配文件数: {len(files)}\n\n---\n\n")

        for rel_path in files:
            abs_path = repo_root / rel_path
            if not abs_path.is_file():
                continue
            if not is_text_extension(rel_path):
                continue
            if is_likely_binary(abs_path):
                continue
            content = read_text_safe(abs_path, MAX_FILE_BYTES)
            if content is None:
                continue
            count += 1
            md.write(f"## {rel_path}\n\n")
            lang = get_lang_hint(rel_path)
            md.write(f"```{lang}\n")
            md.write(content)
            if not content.endswith("\n"):
                md.write("\n")
            md.write("```\n\n")

        md.write(f"\n> ✅ 共导出 {count} 个代码文件。\n")

    print(f"✨ 完成！已生成 {output_path}，包含 {count} 个文件。")

def interactive_mode(repo_root: Path, all_files: list[str]):
    """交互式选择导出范围"""
    print("\n" + "=" * 50)
    print("📦 请选择导出范围：")
    print("  1. 整个仓库")
    print("  2. 指定模块")
    print("  3. 指定目录")
    print("=" * 50)
    choice = input("请输入数字 (1/2/3): ").strip()

    if choice == "1":
        module_filter = None
    elif choice in ("2", "3"):
        filt_type = "模块" if choice == "2" else "目录"
        example = "services/checker" if choice == "2" else "blueprints"
        path = input(f"请输入要导出的{filt_type}路径（相对于仓库根目录，例如 {example}）: ").strip()
        if not path:
            print("路径不能为空，已取消导出。")
            sys.exit(0)
        module_filter = path
    else:
        print("无效的选择，请输入 1、2 或 3。")
        sys.exit(1)

    filtered_files, module_str = filter_files(all_files, module_filter)
    if not filtered_files:
        print(f"警告：在指定条件下未找到任何文件。")
        sys.exit(0)

    print(f"📂 匹配到 {len(filtered_files)} 个文件，开始导出...")
    export_to_markdown(repo_root, filtered_files, module_str)

def main():
    script_dir = Path(__file__).resolve().parent
    repo_root = find_repo_root(script_dir)
    print(f"📂 仓库根目录: {repo_root}")

    all_files = run_git_ls_files(repo_root)
    print(f"📄 仓库未忽略文件总数: {len(all_files)}")

    # 如果提供了命令行参数，则保持非交互模式（兼容自动化）
    if len(sys.argv) > 1:
        # 处理可能的 --help
        if sys.argv[1] in ("-h", "--help"):
            print("用法: python 整合git喂给ai.py [模块路径]")
            print("      无参数时进入交互模式。")
            sys.exit(0)
        module_filter = sys.argv[1].strip()
        filtered_files, module_str = filter_files(all_files, module_filter)
        if not filtered_files:
            print(f"警告：在 '{module_filter}' 下未找到任何文件。")
            sys.exit(0)
        print(f"📂 命令行过滤 '{module_filter}'，匹配 {len(filtered_files)} 个文件。")
        export_to_markdown(repo_root, filtered_files, module_str)
    else:
        interactive_mode(repo_root, all_files)

if __name__ == "__main__":
    main()