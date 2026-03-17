from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.router import api_router
from app.core.config import APP_NAME, APP_VERSION
from app.services.processing_service import warm_processing_cache


app = FastAPI(
    title=APP_NAME,
    version=APP_VERSION,
)


@app.on_event("startup")
def startup_event():
    warm_processing_cache()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # for MVP only
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(api_router)