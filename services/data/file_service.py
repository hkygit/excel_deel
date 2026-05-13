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
        ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
        if ext not in allowed_extensions:
            raise ValueError(f"不支持的文件类型: {ext}")
        filename = secure_filename(file.filename)
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