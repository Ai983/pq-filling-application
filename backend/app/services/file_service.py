import shutil
import uuid
from pathlib import Path

from fastapi import UploadFile

from app.core.config import (
    ALLOWED_EXTENSIONS,
    LOG_DIR,
    PROCESSED_DIR,
    REVIEW_LOG_FILE_SUFFIX,
    UPLOAD_DIR,
)


def validate_file_extension(filename: str) -> bool:
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_EXTENSIONS


def save_uploaded_file(upload_file: UploadFile) -> dict:
    if not upload_file.filename:
        raise ValueError("Uploaded file must have a filename.")

    if not validate_file_extension(upload_file.filename):
        raise ValueError(f"Unsupported file type: {upload_file.filename}")

    file_ext = Path(upload_file.filename).suffix.lower()
    file_id = str(uuid.uuid4())
    stored_name = f"{file_id}{file_ext}"
    stored_path = UPLOAD_DIR / stored_name

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

    with stored_path.open("wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)

    return {
        "file_id": file_id,
        "original_filename": upload_file.filename,
        "stored_filename": stored_name,
        "stored_path": str(stored_path),
    }


def get_uploaded_file_path(file_id: str) -> Path | None:
    if not file_id:
        return None

    for file_path in UPLOAD_DIR.iterdir():
        if file_path.is_file() and file_path.stem == file_id:
            return file_path

    return None


def get_processed_file_path(file_id: str) -> Path:
    return PROCESSED_DIR / f"{file_id}_filled.xlsx"


def get_review_log_path(file_id: str) -> Path:
    return LOG_DIR / f"{file_id}{REVIEW_LOG_FILE_SUFFIX}"