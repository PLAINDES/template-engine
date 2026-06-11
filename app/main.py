# app/main.py
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.core.config import settings
from app.routers import health, docx, sections

app = FastAPI(
    title="DocX Microservice",
    description="Microservicio para parsear y rellenar documentos .docx",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins     = settings.ALLOWED_ORIGINS,
    allow_methods     = ["*"],
    allow_headers     = ["*"],
    allow_credentials = True,
)

app.include_router(health.router)
app.include_router(docx.router)
app.include_router(sections.router)