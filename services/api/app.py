import os
import uuid
import shutil
import json
import yaml
from datetime import datetime
from typing import List, Optional
from pathlib import Path
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form, Request
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from sqlalchemy.orm import Session

from ..database import get_db, engine, Base, User
from ..auth import decode_access_token
from ..user_service import UserService
from .schemas import (
    UserCreate, UserResponse, Token,
    ConfigCreate, ConfigResponse,
    DocumentUploadResponse, FormatRequest, FormatResponse
)

Base.metadata.create_all(bind=engine)

app = FastAPI(
    title="Word文档格式化服务",
    description="Word文档自动化格式处理API",
    version="0.3.0"
)

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token", auto_error=False)

UPLOAD_DIR = "uploads"
GUEST_DIR = os.path.join(UPLOAD_DIR, "guest")
os.makedirs(GUEST_DIR, exist_ok=True)

FRONTEND_DIR = os.path.join(os.path.dirname(__file__), "frontend")
if os.path.exists(FRONTEND_DIR):
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


async def get_optional_user(
    token: Optional[str] = Depends(oauth2_scheme),
    db: Session = Depends(get_db)
) -> Optional[User]:
    if token is None:
        return None
    try:
        payload = decode_access_token(token)
        if payload is None:
            return None
        user_id = payload.get("sub")
        if user_id is None:
            return None
        user_service = UserService(db)
        return user_service.get_user_by_id(int(user_id))
    except Exception:
        return None


async def get_current_user(
    token: str = Depends(oauth2_scheme),
    db: Session = Depends(get_db)
) -> User:
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="无法验证凭据",
        headers={"WWW-Authenticate": "Bearer"},
    )
    if token is None:
        raise credentials_exception
    payload = decode_access_token(token)
    if payload is None:
        raise credentials_exception
    user_id: str = payload.get("sub")
    if user_id is None:
        raise credentials_exception
    user_service = UserService(db)
    user = user_service.get_user_by_id(int(user_id))
    if user is None:
        raise credentials_exception
    return user


@app.get("/", response_class=HTMLResponse, tags=["前端"])
async def index():
    index_path = os.path.join(FRONTEND_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>前端文件未找到</h1>")


@app.post("/register", response_model=UserResponse, tags=["用户认证"])
async def register(user_data: UserCreate, db: Session = Depends(get_db)):
    user_service = UserService(db)
    try:
        user = user_service.register(
            username=user_data.username,
            email=user_data.email,
            password=user_data.password
        )
        return user
    except ValueError as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(e)
        )


@app.post("/token", response_model=Token, tags=["用户认证"])
async def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user_service = UserService(db)
    token = user_service.login(form_data.username, form_data.password)
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="用户名或密码错误",
            headers={"WWW-Authenticate": "Bearer"},
        )
    return {"access_token": token, "token_type": "bearer"}


@app.get("/users/me", response_model=UserResponse, tags=["用户信息"])
async def get_current_user_info(current_user: User = Depends(get_current_user)):
    return current_user


@app.post("/configs", response_model=ConfigResponse, tags=["配置管理"])
async def create_config(
    config_data: ConfigCreate,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    user_service = UserService(db)
    config = user_service.save_user_config(
        user_id=current_user.id,
        config_name=config_data.config_name,
        config_data=config_data.config_data,
        is_default=config_data.is_default
    )
    return config


@app.get("/configs", response_model=List[ConfigResponse], tags=["配置管理"])
async def get_configs(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    user_service = UserService(db)
    return user_service.get_all_configs(current_user.id)


@app.get("/configs/{config_name}", response_model=ConfigResponse, tags=["配置管理"])
async def get_config(
    config_name: str,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    user_service = UserService(db)
    config = user_service.get_user_config(current_user.id, config_name)
    if not config:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="配置不存在"
        )
    return config


@app.get("/api/default-config", tags=["前端接口"])
async def get_default_config():
    config_path = os.path.join("config", "format_config.yaml")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            config_data = yaml.safe_load(f)
        return JSONResponse(content=config_data)
    return JSONResponse(content={})


@app.post("/api/upload", tags=["前端接口"])
async def guest_upload(file: UploadFile = File(...)):
    if not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="仅支持 .docx 格式文件"
        )

    session_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(GUEST_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    safe_filename = file.filename
    file_path = os.path.join(session_dir, safe_filename)

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    file_size = os.path.getsize(file_path)

    return JSONResponse(content={
        "success": True,
        "session_id": session_id,
        "filename": safe_filename,
        "file_size": file_size,
        "message": "文件上传成功"
    })


@app.post("/api/format", tags=["前端接口"])
async def guest_format(
    session_id: str = Form(...),
    filename: str = Form(...),
    config_json: str = Form(default="")
):
    session_dir = os.path.join(GUEST_DIR, session_id)
    input_path = os.path.join(session_dir, filename)

    if not os.path.exists(input_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="文档不存在，请重新上传"
        )

    config_path = os.path.join("config", "format_config.yaml")
    if config_json:
        try:
            user_config = json.loads(config_json)
            config_path = os.path.join(session_dir, "custom_config.yaml")
            with open(config_path, "w", encoding="utf-8") as f:
                yaml.dump(user_config, f, allow_unicode=True, default_flow_style=False)
        except (json.JSONDecodeError, Exception):
            pass

    try:
        sys_path_backup = os.sys.path.copy()
        project_root = str(Path(__file__).parent.parent.parent)
        if project_root not in os.sys.path:
            os.sys.path.insert(0, project_root)

        from main import DocumentFormatter

        if project_root in os.sys.path:
            os.sys.path.remove(project_root)

        formatter = DocumentFormatter(input_path, config_path)
        formatter.format_all()

        base, ext = os.path.splitext(filename)
        output_filename = f"{base}_formatted{ext}"
        output_path = os.path.join(session_dir, output_filename)

        formatter.save(output_path)

        return JSONResponse(content={
            "success": True,
            "session_id": session_id,
            "output_filename": output_filename,
            "message": "文档格式化完成",
            "errors": [err[1] for err in formatter.errors]
        })
    except Exception as e:
        return JSONResponse(content={
            "success": False,
            "message": f"格式化失败: {str(e)}",
            "errors": [str(e)]
        }, status_code=500)


@app.get("/api/download/{session_id}/{filename}", tags=["前端接口"])
async def guest_download(session_id: str, filename: str):
    file_path = os.path.join(GUEST_DIR, session_id, filename)

    if not os.path.exists(file_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="文件不存在"
        )

    display_name = filename.replace("_formatted", "")
    return FileResponse(
        file_path,
        filename=f"formatted_{display_name}",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.post("/documents/upload", response_model=DocumentUploadResponse, tags=["文档处理"])
async def upload_document(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user)
):
    if not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="仅支持 .docx 格式文件"
    )

    user_dir = os.path.join(UPLOAD_DIR, str(current_user.id))
    os.makedirs(user_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{timestamp}_{file.filename}"
    file_path = os.path.join(user_dir, filename)

    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    file_size = os.path.getsize(file_path)

    return DocumentUploadResponse(
        filename=filename,
        file_size=file_size,
        upload_time=datetime.now(),
        message="文件上传成功"
    )


@app.post("/documents/{filename}/format", response_model=FormatResponse, tags=["文档处理"])
async def format_document(
    filename: str,
    format_request: FormatRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    user_dir = os.path.join(UPLOAD_DIR, str(current_user.id))
    input_path = os.path.join(user_dir, filename)

    if not os.path.exists(input_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="文档不存在"
        )

    config_data = format_request.config_data
    if not config_data and format_request.config_name:
        user_service = UserService(db)
        config = user_service.get_user_config(current_user.id, format_request.config_name)
        if config:
            config_data = config.config_data

    if not config_data:
        config_path = "config/format_config.yaml"
    else:
        config_path = os.path.join(user_dir, f"config_{filename}.yaml")
        with open(config_path, "w", encoding="utf-8") as f:
            f.write(config_data)

    try:
        sys_path_backup = os.sys.path.copy()
        project_root = str(Path(__file__).parent.parent.parent)
        if project_root not in os.sys.path:
            os.sys.path.insert(0, project_root)

        from main import DocumentFormatter

        if project_root in os.sys.path:
            os.sys.path.remove(project_root)

        formatter = DocumentFormatter(input_path, config_path)
        formatter.format_all()

        base, ext = os.path.splitext(filename)
        output_filename = f"{base}_formatted{ext}"
        output_path = os.path.join(user_dir, output_filename)

        formatter.save(output_path)

        return FormatResponse(
            success=True,
            output_path=output_filename,
            message="文档格式化完成",
            errors=[err[1] for err in formatter.errors]
        )
    except Exception as e:
        return FormatResponse(
            success=False,
            message=f"格式化失败: {str(e)}",
            errors=[str(e)]
        )


@app.get("/documents/{filename}/download", tags=["文档处理"])
async def download_document(
    filename: str,
    current_user: User = Depends(get_current_user)
):
    user_dir = os.path.join(UPLOAD_DIR, str(current_user.id))
    file_path = os.path.join(user_dir, filename)

    if not os.path.exists(file_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="文档不存在"
        )

    return FileResponse(file_path, filename=filename)


@app.on_event("startup")
async def startup_event():
    from ..database import SessionLocal
    db = SessionLocal()
    try:
        user_service = UserService(db)
        user_service.init_default_roles()
    finally:
        db.close()
