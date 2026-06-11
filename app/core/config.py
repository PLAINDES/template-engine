# app/core/config.py
from pydantic_settings import BaseSettings
from typing import List

class Settings(BaseSettings):
    MINIO_ENDPOINT:   str  = "127.0.0.1"
    MINIO_PORT:       int  = 9100
    MINIO_USE_SSL:    bool = False
    MINIO_ACCESS_KEY: str  = ""
    MINIO_SECRET_KEY: str  = ""
    MINIO_BUCKET:     str  = "prosedi"

    # CORS — en producción pasar los orígenes reales por .env
    ALLOWED_ORIGINS: List[str] = ["*"]
 
    # Validación de archivos
    MAX_FILE_SIZE_MB: int = 20
    ALLOWED_MIME_TYPES: List[str] = [
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/octet-stream",  # algunos clientes usan esto para .docx
    ]
 
    class Config:
        env_file = ".env"
        extra    = "ignore"
 
 
settings = Settings()