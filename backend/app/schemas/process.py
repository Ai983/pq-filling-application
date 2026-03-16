from pydantic import BaseModel


class ProcessResponseData(BaseModel):
    file_id: str
    input_file: str
    filled_file: str
    review_log: str
    total_logged_items: int
    filled_count: int
    skipped_count: int