from werkzeug.security import generate_password_hash, check_password_hash
from typing import Optional, Tuple
import psycopg2


class UserDAO:
    def __init__(self, db):
        self.db = db

    def create_user(self, username: str, password: str) -> Tuple[bool, dict]:
        """Создание нового пользователя"""
        try:
            password_hash = generate_password_hash(password, method='pbkdf2:sha256')
            with self.db.get_cursor() as cur:
                cur.execute(
                    "INSERT INTO users (username, password_hash) VALUES (%s, %s) RETURNING id, username, role",
                    (username, password_hash)
                )
                result = cur.fetchone()
                # Преобразуем результат в словарь
                if result:
                    user_dict = {
                        'id': result[0],
                        'username': result[1],
                        'role': result[2]
                    }
                    return True, user_dict
                else:
                    return False, {'error': 'Не удалось создать пользователя'}

        except psycopg2.IntegrityError as e:
            # Ошибка уникальности (дубликат username)
            return False, {'error': 'Пользователь с таким именем уже существует'}
        except Exception as e:
            return False, {'error': str(e)}

    def get_user_by_username(self, username: str) -> Optional[dict]:
        """Получение пользователя по имени"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, username, password_hash, role FROM users WHERE username = %s", (username,))
            result = cur.fetchone()
            if result:
                return {
                    'id': result[0],
                    'username': result[1],
                    'password_hash': result[2],
                    'role': result[3]
                }
            return None

    def get_user_by_id(self, user_id: int) -> Optional[dict]:
        """Получение пользователя по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, username, role FROM users WHERE id = %s", (user_id,))
            result = cur.fetchone()
            if result:
                return {
                    'id': result[0],
                    'username': result[1],
                    'role': result[2]
                }
            return None

    def verify_password(self, user: dict, password: str) -> bool:
        """Проверка пароля"""
        if user and 'password_hash' in user:
            return check_password_hash(user['password_hash'], password)
        return False
