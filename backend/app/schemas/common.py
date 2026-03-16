from pydantic import BaseModel


class ApiResponse(BaseModel):
    status: str
    message: str
    data: dict | None = None