from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from app.services.file_service import get_processed_file_path, get_review_log_path

router = APIRouter()


@router.get("/download/filled/{file_id}")
def download_filled_file(file_id: str):
    file_path = get_processed_file_path(file_id)

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Filled file not found")

    return FileResponse(
        path=file_path,
        filename=f"{file_id}_filled.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@router.get("/download/log/{file_id}")
def download_review_log(file_id: str):
    file_path = get_review_log_path(file_id)

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Review log not found")

    return FileResponse(
        path=file_path,
        filename=f"{file_id}_review_log.csv",
        media_type="text/csv"
    )