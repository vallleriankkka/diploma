from typing import List, Optional, Dict, Any, Tuple
from datetime import datetime


class ExpertDAO:
    def __init__(self, db):
        self.db = db

    def create(self, expert_data: Dict[str, Any]) -> int:
        """Создание нового эксперта"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO expert (surname, name, patronymic, position, user_id)
                VALUES (%s, %s, %s, %s, %s) RETURNING id
            """, (
                expert_data['surname'],
                expert_data['name'],
                expert_data.get('patronymic', ''),
                expert_data['position'],
                expert_data['user_id']
            ))
            return cur.fetchone()[0]

    def get_by_id(self, expert_id: int) -> Optional[Tuple]:
        """Получение эксперта по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT e.*, u.username 
                FROM expert e
                LEFT JOIN users u ON e.user_id = u.id
                WHERE e.id = %s
            """, (expert_id,))
            return cur.fetchone()

    def get_all(self, user_id: Optional[int] = None) -> List[Tuple]:
        """Получение всех экспертов пользователя"""
        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                    SELECT e.*, 
                           (SELECT COUNT(*) FROM commission_experts ce 
                            WHERE ce.expert_id = e.id) as commission_count
                    FROM expert e
                    WHERE e.user_id = %s
                    ORDER BY e.surname, e.name, e.patronymic
                """, (user_id,))
            else:
                cur.execute("""
                    SELECT e.*, u.username,
                           (SELECT COUNT(*) FROM commission_experts ce 
                            WHERE ce.expert_id = e.id) as commission_count
                    FROM expert e
                    LEFT JOIN users u ON e.user_id = u.id
                    ORDER BY e.surname, e.name, e.patronymic
                """)
            return cur.fetchall()

    def update(self, expert_id: int, expert_data: Dict[str, Any]) -> bool:
        """Обновление данных эксперта"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE expert 
                SET surname = %s, 
                    name = %s, 
                    patronymic = %s, 
                    position = %s
                WHERE id = %s
            """, (
                expert_data['surname'],
                expert_data['name'],
                expert_data.get('patronymic', ''),
                expert_data['position'],
                expert_id
            ))
            return cur.rowcount > 0

    def delete(self, expert_id: int) -> bool:
        """Удаление эксперта (только если он не используется в комиссиях)"""
        with self.db.get_cursor() as cur:
            # Проверяем, используется ли эксперт в комиссиях
            cur.execute("""
                SELECT COUNT(*) FROM commission_experts 
                WHERE expert_id = %s
            """, (expert_id,))

            if cur.fetchone()[0] > 0:
                return False  # Эксперт используется в комиссиях

            cur.execute("DELETE FROM expert WHERE id = %s", (expert_id,))
            return cur.rowcount > 0

    def search(self, search_query: str, user_id: Optional[int] = None) -> List[Tuple]:
        """Поиск экспертов"""
        search_pattern = f"%{search_query}%"

        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                    SELECT e.* 
                    FROM expert e
                    WHERE e.user_id = %s AND (
                        e.surname ILIKE %s OR 
                        e.name ILIKE %s OR 
                        e.patronymic ILIKE %s OR
                        e.position ILIKE %s OR
                        CONCAT(e.surname, ' ', e.name, ' ', COALESCE(e.patronymic, '')) ILIKE %s
                    )
                    ORDER BY e.surname, e.name, e.patronymic
                """, (user_id, search_pattern, search_pattern, search_pattern,
                      search_pattern, search_pattern))
            else:
                cur.execute("""
                    SELECT e.*, u.username
                    FROM expert e
                    LEFT JOIN users u ON e.user_id = u.id
                    WHERE e.surname ILIKE %s OR 
                          e.name ILIKE %s OR 
                          e.patronymic ILIKE %s OR
                          e.position ILIKE %s OR
                          CONCAT(e.surname, ' ', e.name, ' ', COALESCE(e.patronymic, '')) ILIKE %s
                    ORDER BY e.surname, e.name, e.patronymic
                """, (search_pattern, search_pattern, search_pattern,
                      search_pattern, search_pattern))
            return cur.fetchall()

    def is_expert_owner(self, expert_id: int, user_id: int) -> bool:
        """Проверяет, является ли пользователь владельцем эксперта"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) FROM expert 
                WHERE id = %s AND user_id = %s
            """, (expert_id, user_id))
            result = cur.fetchone()
            return result[0] > 0 if result else False
