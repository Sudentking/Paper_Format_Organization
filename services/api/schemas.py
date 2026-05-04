from typing import Optional, List
from datetime import datetime
from pydantic import BaseModel, EmailStr


class UserCreate(BaseModel):
    username: str
    email: EmailStr
    password: str


class UserLogin(BaseModel):
    username: str
    password: str


class UserResponse(BaseModel):
    id: int
    username: str
    email: str
    is_active: bool
    role_id: int
    created_at: datetime

    class Config:
        from_attributes = True


class Token(BaseModel):
    access_token: str
    token_type: str


class ConfigCreate(BaseModel):
    config_name: str
    config_data: str
    is_default: bool = False


class ConfigResponse(BaseModel):
    id: int
    user_id: int
    config_name: str
    config_data: str
    is_default: bool
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


class DocumentUploadResponse(BaseModel):
    filename: str
    file_size: int
    upload_time: datetime
    message: str


class FormatRequest(BaseModel):
    config_name: Optional[str] = None
    config_data: Optional[str] = None


class FormatResponse(BaseModel):
    success: bool
    output_path: Optional[str] = None
    message: str
    errors: List[str] = []
