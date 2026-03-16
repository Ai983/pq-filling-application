from fastapi import APIRouter
from app.api.routes.health import router as health_router
from app.api.routes.upload import router as upload_router
from app.api.routes.process import router as process_router
from app.api.routes.download import router as download_router

api_router = APIRouter()
api_router.include_router(health_router, tags=["Health"])
api_router.include_router(upload_router, tags=["Upload"])
api_router.include_router(process_router, tags=["Process"])
api_router.include_router(download_router, tags=["Download"])