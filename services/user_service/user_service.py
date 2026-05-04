from typing import Optional, List
from sqlalchemy.orm import Session
from ..database.models import User, Role, UserConfig
from ..auth import verify_password, get_password_hash, create_access_token


class UserService:
    def __init__(self, db: Session):
        self.db = db

    def register(self, username: str, email: str, password: str) -> User:
        existing_user = self.db.query(User).filter(
            (User.username == username) | (User.email == email)
        ).first()
        if existing_user:
            raise ValueError("用户名或邮箱已存在")

        hashed_password = get_password_hash(password)
        user = User(
            username=username,
            email=email,
            hashed_password=hashed_password,
            role_id=2
        )
        self.db.add(user)
        self.db.commit()
        self.db.refresh(user)
        return user

    def login(self, username: str, password: str) -> Optional[str]:
        user = self.db.query(User).filter(User.username == username).first()
        if not user or not verify_password(password, user.hashed_password):
            return None
        if not user.is_active:
            return None

        access_token = create_access_token(data={"sub": str(user.id)})
        return access_token

    def get_user_by_id(self, user_id: int) -> Optional[User]:
        return self.db.query(User).filter(User.id == user_id).first()

    def get_user_by_username(self, username: str) -> Optional[User]:
        return self.db.query(User).filter(User.username == username).first()

    def update_user(self, user_id: int, **kwargs) -> Optional[User]:
        user = self.get_user_by_id(user_id)
        if not user:
            return None
        for key, value in kwargs.items():
            if hasattr(user, key) and key != 'id':
                setattr(user, key, value)
        self.db.commit()
        self.db.refresh(user)
        return user

    def deactivate_user(self, user_id: int) -> bool:
        user = self.get_user_by_id(user_id)
        if not user:
            return False
        user.is_active = False
        self.db.commit()
        return True

    def get_user_config(self, user_id: int, config_name: Optional[str] = None) -> Optional[UserConfig]:
        query = self.db.query(UserConfig).filter(UserConfig.user_id == user_id)
        if config_name:
            query = query.filter(UserConfig.config_name == config_name)
        else:
            query = query.filter(UserConfig.is_default == True)
        return query.first()

    def save_user_config(self, user_id: int, config_name: str, config_data: str, is_default: bool = False) -> UserConfig:
        existing = self.db.query(UserConfig).filter(
            UserConfig.user_id == user_id,
            UserConfig.config_name == config_name
        ).first()

        if existing:
            existing.config_data = config_data
            existing.is_default = is_default
            self.db.commit()
            self.db.refresh(existing)
            return existing

        config = UserConfig(
            user_id=user_id,
            config_name=config_name,
            config_data=config_data,
            is_default=is_default
        )
        self.db.add(config)
        self.db.commit()
        self.db.refresh(config)
        return config

    def get_all_configs(self, user_id: int) -> List[UserConfig]:
        return self.db.query(UserConfig).filter(UserConfig.user_id == user_id).all()

    def init_default_roles(self):
        default_roles = [
            {"name": "admin", "description": "系统管理员"},
            {"name": "user", "description": "普通用户"},
            {"name": "guest", "description": "访客"},
        ]
        for role_data in default_roles:
            existing = self.db.query(Role).filter(Role.name == role_data["name"]).first()
            if not existing:
                role = Role(**role_data)
                self.db.add(role)
        self.db.commit()
