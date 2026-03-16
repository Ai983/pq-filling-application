from fastapi import APIRouter, File, HTTPException, UploadFile

from app.services.file_service import save_uploaded_file, validate_file_extension

router = APIRouter()


@router.post("/upload")
def upload_excel(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file selected")

    if not validate_file_extension(file.filename):
        raise HTTPException(status_code=400, detail="Only .xlsx files are allowed")

    try:
        result = save_uploaded_file(file)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded file: {str(exc)}")

    return {
        "status": "success",
        "message": "File uploaded successfully",
        "data": result,
    }