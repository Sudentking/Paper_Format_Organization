import sys
import os
import unittest
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from services.database.database import Base
from services.database.models import User, Role, UserConfig
from services.user_service.user_service import UserService
from services.auth import get_password_hash, verify_password, create_access_token, decode_access_token


class TestAuthHandler(unittest.TestCase):
    def test_password_hash(self):
        password = "testpassword123"
        hashed = get_password_hash(password)
        self.assertNotEqual(password, hashed)
        self.assertTrue(verify_password(password, hashed))

    def test_password_verify_wrong(self):
        password = "testpassword123"
        hashed = get_password_hash(password)
        self.assertFalse(verify_password("wrongpassword", hashed))

    def test_create_access_token(self):
        data = {"sub": "123"}
        token = create_access_token(data)
        self.assertIsNotNone(token)
        self.assertIsInstance(token, str)

    def test_decode_access_token(self):
        data = {"sub": "123"}
        token = create_access_token(data)
        payload = decode_access_token(token)
        self.assertIsNotNone(payload)
        self.assertEqual(payload["sub"], "123")

    def test_decode_invalid_token(self):
        payload = decode_access_token("invalid.token.here")
        self.assertIsNone(payload)


class TestUserService(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.db = self.Session()
        self.service = UserService(self.db)

    def tearDown(self):
        self.db.close()
        Base.metadata.drop_all(self.engine)

    def test_register_user(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        self.assertIsNotNone(user)
        self.assertEqual(user.username, "testuser")
        self.assertEqual(user.email, "test@example.com")
        self.assertTrue(user.is_active)
        self.assertEqual(user.role_id, 2)

    def test_register_duplicate_username(self):
        self.service.register("testuser", "test1@example.com", "password123")
        with self.assertRaises(ValueError):
            self.service.register("testuser", "test2@example.com", "password456")

    def test_register_duplicate_email(self):
        self.service.register("testuser1", "test@example.com", "password123")
        with self.assertRaises(ValueError):
            self.service.register("testuser2", "test@example.com", "password456")

    def test_login_success(self):
        self.service.register("testuser", "test@example.com", "password123")
        token = self.service.login("testuser", "password123")
        self.assertIsNotNone(token)

    def test_login_wrong_password(self):
        self.service.register("testuser", "test@example.com", "password123")
        token = self.service.login("testuser", "wrongpassword")
        self.assertIsNone(token)

    def test_login_nonexistent_user(self):
        token = self.service.login("nonexistent", "password123")
        self.assertIsNone(token)

    def test_get_user_by_id(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        found = self.service.get_user_by_id(user.id)
        self.assertIsNotNone(found)
        self.assertEqual(found.username, "testuser")

    def test_get_user_by_username(self):
        self.service.register("testuser", "test@example.com", "password123")
        found = self.service.get_user_by_username("testuser")
        self.assertIsNotNone(found)
        self.assertEqual(found.email, "test@example.com")

    def test_deactivate_user(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        result = self.service.deactivate_user(user.id)
        self.assertTrue(result)
        updated_user = self.service.get_user_by_id(user.id)
        self.assertFalse(updated_user.is_active)

    def test_save_user_config(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        config = self.service.save_user_config(
            user.id, "default", '{"font": "宋体"}', is_default=True
        )
        self.assertIsNotNone(config)
        self.assertEqual(config.config_name, "default")
        self.assertTrue(config.is_default)

    def test_get_user_config(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        self.service.save_user_config(user.id, "default", '{"font": "宋体"}', is_default=True)
        config = self.service.get_user_config(user.id, "default")
        self.assertIsNotNone(config)
        self.assertEqual(config.config_data, '{"font": "宋体"}')

    def test_get_all_configs(self):
        user = self.service.register("testuser", "test@example.com", "password123")
        self.service.save_user_config(user.id, "config1", '{"font": "宋体"}')
        self.service.save_user_config(user.id, "config2", '{"font": "黑体"}')
        configs = self.service.get_all_configs(user.id)
        self.assertEqual(len(configs), 2)

    def test_init_default_roles(self):
        self.service.init_default_roles()
        admin_role = self.db.query(Role).filter(Role.name == "admin").first()
        user_role = self.db.query(Role).filter(Role.name == "user").first()
        guest_role = self.db.query(Role).filter(Role.name == "guest").first()
        self.assertIsNotNone(admin_role)
        self.assertIsNotNone(user_role)
        self.assertIsNotNone(guest_role)


if __name__ == '__main__':
    unittest.main()
