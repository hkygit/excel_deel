#!/usr/bin/env python3
"""
一键导出 Git 仓库中未被 .gitignore 忽略的代码文件为 POML (Project Object Markup Language)
支持模式：
1. 整个仓库（全量导出）
2. 根据配置文件导出（从 POML 目录读取文件列表，每行一个文件或目录，支持 # 注释）
"""

import subprocess
import os
import sys
from pathlib import Path

# ---------------------- 可自定义常量 ----------------------
TEXT_EXTENSIONS = {
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
    ".js", ".mjs", ".cjs", ".ts", ".jsx", ".tsx",
    ".vue", ".svelte", ".astro",
    ".html", ".htm", ".xhtml",
    ".css", ".scss", ".sass", ".less", ".styl",
    ".xml", ".xsl", ".xslt", ".svg",
    ".json", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".conf", ".properties",
    ".env", ".env.example", ".editorconfig", ".gitignore", ".gitattributes",
    ".md", ".markdown", ".rst", ".txt", ".csv", ".tsv",
    ".sql", ".graphql", ".gql", ".proto", ".avsc",
}
MAX_FILE_BYTES = 500 * 1024          # 500 KB
POML_DIR = "POML"                    # 输出根目录下的文件夹名
OUTPUT_FILENAME_FULL = "full_repository.md"
# ----------------------------------------------------------

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
    """调用 git ls-files 返回所有未被忽略的文件（已跟踪或未跟踪但未忽略）"""
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
    try:
        with open(filepath, "rb") as f:
            chunk = f.read(1024)
        return b"\x00" in chunk
    except Exception:
        return True

def read_text_safe(filepath: Path, max_bytes: int) -> str | None:
    try:
        size = filepath.stat().st_size
        if size > max_bytes:
            return None
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except Exception:
        return None

def get_lang_hint(filename: str) -> str:
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

def expand_paths(rel_paths: list[str], all_tracked_files: list[str]) -> list[str]:
    """
    将路径列表（可能包含目录）展开为实际文件列表。
    - 文件：直接加入
    - 目录：找出所有以该目录开头的文件（前缀匹配）
    """
    result = set()
    for p in rel_paths:
        p = p.strip("/\\")
        if not p:
            continue
        if p in all_tracked_files:
            result.add(p)
        else:
            prefix = p + "/"
            matched = [f for f in all_tracked_files if f.startswith(prefix)]
            if matched:
                result.update(matched)
            else:
                print(f"⚠️ 警告: 路径不存在或不在 Git 跟踪中，已忽略: {p}", file=sys.stderr)
    return sorted(result)

def load_config_file(config_name: str, repo_root: Path) -> list[str]:
    """读取配置文件（位于 POML_DIR 目录下），返回非空、非注释的行列表"""
    config_path = repo_root / POML_DIR / config_name
    if not config_path.exists():
        print(f"错误: 配置文件不存在: {config_path}", file=sys.stderr)
        sys.exit(1)
    lines = []
    with open(config_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            lines.append(line)
    return lines

def export_poml(repo_root: Path, files: list[str], output_name: str):
    """
    导出为 POML 格式（结构化 Markdown 列表）
    output_name: 输出文件名（不含路径），将保存在 POML_DIR 下
    """
    out_dir = repo_root / POML_DIR
    out_dir.mkdir(exist_ok=True)
    output_path = out_dir / output_name

    count = 0
    with open(output_path, "w", encoding="utf-8") as md:
        md.write(f"# Project Object Markup Language (POML)\n\n")
        md.write(f"- **repository**: `{repo_root.name}`\n")
        md.write(f"- **export_filter**: `{output_name.replace('.md', '')}`\n")
        md.write(f"- **total_files**: {len(files)}\n\n")
        md.write(f"## Files\n\n")

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
            lang = get_lang_hint(rel_path)
            md.write(f"- **{rel_path}**\n")
            md.write(f"  - language: `{lang}`\n")
            md.write(f"  - content:\n")
            md.write(f"    ```{lang}\n")
            md.write(content)
            if not content.endswith("\n"):
                md.write("\n")
            md.write(f"    ```\n\n")

        md.write(f"\n> ✅ 共导出 {count} 个代码文件。\n")

    print(f"✨ 完成！已生成 {output_path}，包含 {count} 个文件。")

def interactive_mode(repo_root: Path, all_files: list[str]):
    """交互式选择导出范围"""
    print("\n" + "=" * 60)
    print("📦 请选择导出范围：")
    print("  1. 整个仓库（全量导出）")
    print("  2. 根据配置文件导出（从 POML 目录读取文件列表）")
    print("=" * 60)
    choice = input("请输入数字 (1/2): ").strip()

    if choice == "1":
        print(f"📂 匹配到 {len(all_files)} 个文件，开始导出...")
        export_poml(repo_root, all_files, OUTPUT_FILENAME_FULL)
    elif choice == "2":
        config_name = input("请输入配置文件名（位于 POML/ 目录下，如 checker.txt）: ").strip()
        if not config_name:
            print("文件名不能为空，已取消导出。")
            sys.exit(0)
        rel_paths = load_config_file(config_name, repo_root)
        if not rel_paths:
            print("配置文件为空或无有效行，已取消导出。")
            sys.exit(0)
        target_files = expand_paths(rel_paths, all_files)
        if not target_files:
            print("未找到任何有效文件，请检查配置文件中的路径是否正确。")
            sys.exit(0)
        print(f"📂 配置文件指定了 {len(rel_paths)} 个条目，展开后得到 {len(target_files)} 个文件。")
        base_name = os.path.splitext(config_name)[0]
        output_name = f"{base_name}.md"
        export_poml(repo_root, target_files, output_name)
    else:
        print("无效的选择，请输入 1 或 2。")
        sys.exit(1)

def main():
    script_dir = Path(__file__).resolve().parent
    repo_root = find_repo_root(script_dir)
    print(f"📂 仓库根目录: {repo_root}")

    all_files = run_git_ls_files(repo_root)
    print(f"📄 仓库未忽略文件总数: {len(all_files)}")

    # 命令行参数兼容
    if len(sys.argv) > 1:
        if sys.argv[1] in ("-h", "--help"):
            print("用法: python 整合git喂给ai.py           # 交互模式")
            print("       python 整合git喂给ai.py --help  # 显示帮助")
            sys.exit(0)
        else:
            print("警告：命令行直接指定模块的功能已移除，请使用交互模式。")
            interactive_mode(repo_root, all_files)
    else:
        interactive_mode(repo_root, all_files)

if __name__ == "__main__":
    main()