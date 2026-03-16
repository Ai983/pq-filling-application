from pydantic import BaseModel


class UploadResponseData(BaseModel):
    file_id: str
    original_filename: str
    stored_filename: str
    stored_path: str