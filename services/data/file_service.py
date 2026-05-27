# utils/services/file_service.py
import os
import uuid
import logging
import tempfile
from typing import List, Optional
from werkzeug.utils import secure_filename
from flask import current_app

logger = logging.getLogger(__name__)

class FileService:
    """统一文件管理服务"""

    @staticmethod
    def save_upload(file, subdir: str = '', allowed_extensions: Optional[List[str]] = None) -> str:
        if not file or not file.filename:
            raise ValueError("无效文件")
        if allowed_extensions is None:
            allowed_extensions = current_app.config.get('ALLOWED_EXTENSIONS', {'xlsx', 'xls', 'csv'})

        # 分离文件名和扩展名，保留原始扩展名
        original_filename = file.filename
        name, ext = os.path.splitext(original_filename)
        # 对文件名部分进行安全处理，扩展名保留原样（不经过 secure_filename，避免点号被替换）
        safe_name = secure_filename(name) if name else 'file'
        if not safe_name:
            safe_name = 'file'
        # 确保扩展名以小写形式保存
        ext_lower = ext.lower()
        # 如果扩展名不在允许列表中，但文件实际是有效的 Excel，我们仍保留（后续会转换）
        # 这里不做严格检查，交给后续处理
        filename = f"{safe_name}{ext_lower}"

        unique_id = uuid.uuid4().hex
        base_dir = current_app.config.get('UPLOAD_FOLDER', 'uploads')
        save_dir = os.path.join(base_dir, subdir) if subdir else base_dir
        os.makedirs(save_dir, exist_ok=True)
        save_path = os.path.join(save_dir, f"{unique_id}_{filename}")
        file.save(save_path)
        logger.debug(f"文件已保存: {save_path}")
        return save_path

    @staticmethod
    def cleanup(paths: List[str]):
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    logger.debug(f"已删除文件: {path}")
            except Exception as e:
                logger.warning(f"删除文件失败 {path}: {e}")

    @staticmethod
    def get_temp_path(suffix: str = '') -> str:
        fd, path = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        return path

    @staticmethod
    def safe_filename(name: str) -> str:
        return secure_filename(name)
