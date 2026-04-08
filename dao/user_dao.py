from werkzeug.security import generate_password_hash, check_password_hash
from typing import Optional, Tuple, List, Dict
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
            return False, {'error': 'Пользователь с таким именем уже существует'}
        except Exception as e:
            return False, {'error': str(e)}

    def get_all_users(self) -> List[Dict]:
        """Получение всех пользователей"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT id, username, role, created_at 
                FROM users 
                ORDER BY id
            """)
            results = cur.fetchall()
            users = []
            for row in results:
                users.append({
                    'id': row[0],
                    'username': row[1],
                    'role': row[2],
                    'created_at': row[3]
                })
            return users

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

    def update_user_role(self, user_id: int, role: str) -> bool:
        """Обновление роли пользователя"""
        if role not in ['user', 'admin']:
            return False

        with self.db.get_cursor() as cur:
            cur.execute("UPDATE users SET role = %s WHERE id = %s", (role, user_id))
            return cur.rowcount > 0

    def delete_user(self, user_id: int) -> Tuple[bool, str]:
        """Удаление пользователя"""
        # Проверяем, не является ли пользователь единственным админом
        with self.db.get_cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
            admin_count = cur.fetchone()[0]

            if admin_count == 1:
                cur.execute("SELECT role FROM users WHERE id = %s", (user_id,))
                user_role = cur.fetchone()
                if user_role and user_role[0] == 'admin':
                    return False, "Нельзя удалить единственного администратора"

        with self.db.get_cursor() as cur:
            cur.execute("DELETE FROM users WHERE id = %s", (user_id,))
            if cur.rowcount > 0:
                return True, "Пользователь успешно удален"
            else:
                return False, "Пользователь не найден"

    def create_user_by_admin(self, username: str, password: str, role: str = 'user') -> Tuple[bool, dict]:
        """Создание нового пользователя администратором"""
        try:
            password_hash = generate_password_hash(password, method='pbkdf2:sha256')
            with self.db.get_cursor() as cur:
                cur.execute(
                    "INSERT INTO users (username, password_hash, role) VALUES (%s, %s, %s) RETURNING id, username, role",
                    (username, password_hash, role)
                )
                result = cur.fetchone()
                if result:
                    user_dict = {
                        'id': result[0],
                        'username': result[1],
                        'role': result[2]
                    }
                    return True, user_dict
                else:
                    return False, {'error': 'Не удалось создать пользователя'}

        except psycopg2.IntegrityError:
            return False, {'error': 'Пользователь с таким именем уже существует'}
        except Exception as e:
            return False, {'error': str(e)}
