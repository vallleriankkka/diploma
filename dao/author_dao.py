from typing import List, Optional
from datetime import date, datetime


class AuthorDAO:
    def __init__(self, db):
        self.db = db

    def get_all(self, user_id=None, with_education_info=False):
        with self.db.get_cursor() as cur:
            if user_id:
                # Пользователь - только своих авторов
                if with_education_info:
                    cur.execute("""
                        SELECT a.*, el.name as education_level_name
                        FROM author a 
                        LEFT JOIN education_level el ON a.education_level_id = el.id
                        WHERE a.user_id = %s 
                        ORDER BY a.surname, a.name
                    """, (user_id,))
                else:
                    cur.execute("""
                        SELECT a.* 
                        FROM author a 
                        WHERE a.user_id = %s 
                        ORDER BY a.surname, a.name
                    """, (user_id,))
            else:
                # Админ - всех авторов
                if with_education_info:
                    cur.execute("""
                        SELECT a.*, u.username, el.name as education_level_name
                        FROM author a 
                        LEFT JOIN users u ON a.user_id = u.id
                        LEFT JOIN education_level el ON a.education_level_id = el.id
                        ORDER BY a.surname, a.name
                    """)
                else:
                    cur.execute("""
                        SELECT a.*, u.username 
                        FROM author a 
                        LEFT JOIN users u ON a.user_id = u.id
                        ORDER BY a.surname, a.name
                    """)
            return cur.fetchall()

    def get_by_id(self, author_id: int) -> Optional[tuple]:
        """Получение автора по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT a.*, el.name as education_level_name
                FROM author a 
                LEFT JOIN education_level el ON a.education_level_id = el.id
                WHERE a.id = %s
            """, (author_id,))
            return cur.fetchone()

    def create(self, author_data: dict) -> int:
        """Создание нового автора"""
        with self.db.get_cursor() as cur:
            # Если автор не студент, обнуляем education_level_id
            education_level_id = author_data.get('education_level_id') if author_data.get('is_student', False) else None

            cur.execute("""
                INSERT INTO author (
                    surname, name, patronymic, birth_date, address, 
                    department, passport, passport_issue_date, passport_issued_by, passport_issued_code,
                    snils, inn, is_student, position, user_id, citizenship,
                    education_level_id
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (
                author_data['surname'],
                author_data['name'],
                author_data.get('patronymic', ''),
                author_data['birth_date'],
                author_data['address'],
                author_data['department'],
                author_data['passport'],
                author_data.get('passport_issue_date'),
                author_data.get('passport_issued_by', ''),
                author_data.get('passport_issued_code', ''),
                author_data['snils'],
                author_data['inn'],
                author_data.get('is_student', False),
                author_data.get('position', ''),
                author_data['user_id'],
                author_data.get('citizenship', 'Российская Федерация'),
                education_level_id
            ))
            return cur.fetchone()[0]

    def update(self, author_id: int, author_data: dict) -> bool:
        """Обновление автора"""
        with self.db.get_cursor() as cur:
            # Сначала получаем текущее состояние автора
            cur.execute("SELECT is_student FROM author WHERE id = %s", (author_id,))
            current = cur.fetchone()
            current_is_student = current[0] if current else False

            # Определяем is_student из данных или оставляем текущее
            if 'is_student' in author_data:
                is_student = author_data.get('is_student')
                if isinstance(is_student, str):
                    is_student = is_student.lower() in ('true', '1', 'yes', 'on', 'да')
                else:
                    is_student = bool(is_student)
            else:
                is_student = current_is_student

            # Обрабатываем education_level_id
            if is_student:
                education_level_id = author_data.get('education_level_id')
                # Можно добавить проверку, что уровень существует
                if education_level_id:
                    cur.execute("SELECT id FROM education_level WHERE id = %s", (education_level_id,))
                    if not cur.fetchone():
                        education_level_id = None
            else:
                education_level_id = None

            cur.execute("""
                UPDATE author SET 
                    surname=%s, 
                    name=%s, 
                    patronymic=%s, 
                    birth_date=%s, 
                    address=%s, 
                    department=%s, 
                    passport=%s, 
                    passport_issue_date=%s,
                    passport_issued_by=%s,
                    passport_issued_code=%s,
                    snils=%s, 
                    inn=%s, 
                    is_student=%s, 
                    position=%s, 
                    citizenship=%s,
                    education_level_id=%s
                WHERE id=%s
            """, (
                author_data['surname'],
                author_data['name'],
                author_data.get('patronymic', ''),
                author_data['birth_date'],
                author_data['address'],
                author_data['department'],
                author_data['passport'],
                author_data.get('passport_issue_date'),
                author_data.get('passport_issued_by', ''),
                author_data.get('passport_issued_code', ''),
                author_data['snils'],
                author_data['inn'],
                is_student,  # Используем обработанное значение
                author_data.get('position', ''),
                author_data.get('citizenship', 'Российская Федерация'),
                education_level_id,
                author_id
            ))
            return cur.rowcount > 0

    def delete(self, author_id: int) -> bool:
        """Удаление автора"""
        with self.db.get_cursor() as cur:
            cur.execute("DELETE FROM author WHERE id = %s", (author_id,))
            return cur.rowcount > 0

    def check_unique(self, field_name: str, value: str, exclude_id: Optional[int] = None) -> bool:
        """Проверка уникальности поля"""
        with self.db.get_cursor() as cur:
            if exclude_id:
                cur.execute(f"SELECT id FROM author WHERE {field_name} = %s AND id != %s",
                            (value, exclude_id))
            else:
                cur.execute(f"SELECT id FROM author WHERE {field_name} = %s", (value,))
            return cur.fetchone() is not None

    def get_all_education_levels(self):
        """Получение всех уровней образования для выпадающего списка"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, name FROM education_level ORDER BY name")
            return cur.fetchall()

    def get_by_user(self, author_id: int, user_id: int) -> Optional[int]:
        """Проверка принадлежности автора пользователю"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT user_id FROM author WHERE id = %s", (author_id,))
            result = cur.fetchone()
            return result[0] if result else None

