from fastapi import APIRouter, HTTPException

from app.services.processing_service import process_uploaded_file

router = APIRouter()


@router.post("/process/{file_id}")
def process_file(file_id: str):
    try:
        result = process_uploaded_file(file_id)
        return {
            "status": "success",
            "message": "File processed successfully",
            "data": result,
        }
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(exc)}")