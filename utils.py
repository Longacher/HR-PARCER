import os
import uuid
import time

def get_unique_filename(prefix: str, original_name: str, default_ext: str) -> str:
    """
    Генерирует уникальное имя файла с сохранением расширения.
    """
    _, ext = os.path.splitext(original_name)
    if not ext:
        ext = default_ext
    return f"{prefix}_{uuid.uuid4()}{ext}"

def wait_for_new_file(download_dir: str, existing_files: set, timeout_sec: int = 20) -> str:
    """
    Ожидает появления нового файла в указанной директории, отличного от списка existing_files.
    """
    end_time = time.time() + timeout_sec
    while time.time() < end_time:
        current_files = set(os.listdir(download_dir))
        diff = current_files - existing_files
        for candidate in diff:
            if not candidate.endswith(".crdownload"):
                return candidate
        time.sleep(0.5)
    return None
