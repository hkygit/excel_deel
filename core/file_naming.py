# core/file_naming.py
import os
from urllib.parse import quote

def generate_output_filename(module_name: str, original_filename: str, suffix: str = '_小计') -> str:
    base = os.path.splitext(os.path.basename(original_filename))[0]
    safe_module = module_name.replace('/', '_').replace('\\', '_')
    return f"{safe_module}_{base}{suffix}.xlsx"


def build_content_disposition(filename: str) -> str:
    encoded = quote(filename, safe='')
    fallback = ''.join(ch if ord(ch) < 128 else '_' for ch in filename)
    if not fallback:
        fallback = 'download.xlsx'
    return f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{encoded}"
