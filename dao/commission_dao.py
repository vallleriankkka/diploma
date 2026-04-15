from typing import List, Optional, Dict, Any, Tuple
from datetime import datetime


class CommissionDAO:
    def __init__(self, db):
        self.db = db

    def get_by_id(self, commission_id: int) -> Optional[Tuple]:
        """Получение комиссии по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT ec.*, u.username,
                       (SELECT COUNT(*) FROM commission_experts ce 
                        WHERE ce.commission_id = ec.id) as member_count
                FROM commissions ec
                LEFT JOIN users u ON ec.user_id = u.id
                WHERE ec.id = %s
            """, (commission_id,))
            return cur.fetchone()

    def get_all(self, user_id: Optional[int] = None) -> List[Tuple]:
        """Получение всех комиссий пользователя"""
        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                    SELECT ec.*,
                           (SELECT COUNT(*) FROM commission_experts ce 
                            WHERE ce.commission_id = ec.id) as member_count
                    FROM commissions ec
                    WHERE ec.user_id = %s
                    ORDER BY ec.created_date DESC
                """, (user_id,))
            else:
                cur.execute("""
                    SELECT ec.*, u.username,
                           (SELECT COUNT(*) FROM commission_experts ce 
                            WHERE ce.commission_id = ec.id) as member_count
                    FROM commissions ec
                    LEFT JOIN users u ON ec.user_id = u.id
                    ORDER BY ec.created_date DESC
                """)
            return cur.fetchall()

    def create(self, name: str, description: str, user_id: int) -> int:
        """Создание новой комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO commissions (name, description, user_id, created_date)
                VALUES (%s, %s, %s, %s) RETURNING id
            """, (name, description, user_id, datetime.now()))
            return cur.fetchone()[0]

    def update(self, commission_id: int, name: str, description: str) -> bool:
        """Обновление комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE commissions 
                SET name = %s, description = %s
                WHERE id = %s
            """, (name, description, commission_id))
            return cur.rowcount > 0

    def delete(self, commission_id: int) -> bool:
        """Удаление комиссии"""
        with self.db.get_cursor() as cur:
            # Сначала удаляем связи с экспертами
            cur.execute("DELETE FROM commission_experts WHERE commission_id = %s",
                        (commission_id,))
            # Затем удаляем комиссию
            cur.execute("DELETE FROM commissions WHERE id = %s", (commission_id,))
            return cur.rowcount > 0

    def add_expert_to_commission(self, commission_id: int, expert_id: int, role_id: int) -> bool:
        """Добавление существующего эксперта в комиссию"""
        try:
            with self.db.get_cursor() as cur:
                # Проверяем, не состоит ли уже эксперт в комиссии
                cur.execute("""
                    SELECT COUNT(*) FROM commission_experts 
                    WHERE commission_id = %s AND expert_id = %s
                """, (commission_id, expert_id))

                if cur.fetchone()[0] > 0:
                    return False  # Эксперт уже в комиссии

                cur.execute("""
                    INSERT INTO commission_experts (commission_id, expert_id, role_id)
                    VALUES (%s, %s, %s)
                """, (commission_id, expert_id, role_id))
                return True
        except Exception as e:
            print(f"Error adding expert to commission: {e}")
            return False

    def remove_expert_from_commission(self, commission_id: int, expert_id: int) -> bool:
        """Удаление эксперта из комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                DELETE FROM commission_experts 
                WHERE commission_id = %s AND expert_id = %s
            """, (commission_id, expert_id))
            return cur.rowcount > 0

    def update_expert_role(self, commission_id: int, expert_id: int, role_id: int) -> bool:
        """Обновление роли эксперта в комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE commission_experts 
                SET role_id = %s 
                WHERE commission_id = %s AND expert_id = %s
            """, (role_id, commission_id, expert_id))
            return cur.rowcount > 0

    def get_commission_experts(self, commission_id: int) -> List[Tuple]:
        """Получение экспертов комиссии"""
        with self.db.get_cursor() as cur:
            try:
                # Пытаемся получить с description
                cur.execute("""
                    SELECT ce.*, e.surname, e.name, e.patronymic, e.position,
                           cr.name as role_name
                    FROM commission_experts ce
                    JOIN expert e ON ce.expert_id = e.id
                    JOIN commission_roles cr ON ce.role_id = cr.id
                    WHERE ce.commission_id = %s
                    ORDER BY 
                        CASE cr.name
                            WHEN 'председатель' THEN 1
                            WHEN 'эксперт' THEN 2
                            WHEN 'секретарь' THEN 3
                            ELSE 4
                        END,
                        e.surname, e.name, e.patronymic
                """, (commission_id,))
                return cur.fetchall()
            except:
                # Если ошибка, получаем без description
                cur.execute("""
                    SELECT ce.*, e.surname, e.name, e.patronymic, e.position,
                           cr.name as role_name
                    FROM commission_experts ce
                    JOIN expert e ON ce.expert_id = e.id
                    JOIN commission_roles cr ON ce.role_id = cr.id
                    WHERE ce.commission_id = %s
                    ORDER BY 
                        CASE cr.name
                            WHEN 'председатель' THEN 1
                            WHEN 'эксперт' THEN 2
                            WHEN 'секретарь' THEN 3
                            ELSE 4
                        END,
                        e.surname, e.name, e.patronymic
                """, (commission_id,))
                return cur.fetchall()

    def get_available_experts_for_commission(self, commission_id: int, user_id: int) -> List[Tuple]:
        """Получение доступных экспертов (не входящих в комиссию) для пользователя"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT e.* 
                FROM expert e
                WHERE e.user_id = %s 
                AND e.id NOT IN (
                    SELECT expert_id FROM commission_experts 
                    WHERE commission_id = %s
                )
                ORDER BY e.surname, e.name, e.patronymic
            """, (user_id, commission_id))
            return cur.fetchall()

    def get_commission_roles(self) -> List[Tuple]:
        """Получение всех ролей"""
        with self.db.get_cursor() as cur:
            try:
                cur.execute("SELECT * FROM commission_roles ORDER BY id")
                return cur.fetchall()
            except:
                # Если таблицы нет, возвращаем пустой список
                return []

    def has_chairman(self, commission_id: int) -> bool:
        """Проверка, есть ли в комиссии председатель"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) 
                FROM commission_experts ce
                JOIN commission_roles cr ON ce.role_id = cr.id
                WHERE ce.commission_id = %s 
                AND cr.name = 'председатель'
            """, (commission_id,))
            result = cur.fetchone()
            return result[0] > 0 if result else False

    def is_commission_owner(self, commission_id: int, user_id: int) -> bool:
        """Проверяет, является ли пользователь владельцем комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) FROM commissions 
                WHERE id = %s AND user_id = %s
            """, (commission_id, user_id))
            result = cur.fetchone()
            return result[0] > 0 if result else False

    def get_by_user(self, commission_id: int, user_id: int) -> Optional[int]:
        """Проверка принадлежности комиссии пользователю"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT user_id FROM commissions WHERE id = %s", (commission_id,))
            result = cur.fetchone()
            return result[0] if result else None

    def get_commission_experts_detailed(self, commission_id):
        """Получение экспертов комиссии с полными данными"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT 
                    e.id, 
                    e.surname, 
                    e.name, 
                    e.patronymic, 
                    e.position,
                    cr.name as role_name,
                    e.user_id
                FROM commission_experts ce
                JOIN expert e ON ce.expert_id = e.id
                JOIN commission_roles cr ON ce.role_id = cr.id
                WHERE ce.commission_id = %s
                ORDER BY 
                    CASE cr.name
                        WHEN 'председатель' THEN 1
                        WHEN 'эксперт' THEN 2
                        ELSE 3
                    END,
                    e.surname, e.name, e.patronymic
            """, (commission_id,))
            return cur.fetchall()

    def get_commission_by_id_with_user(self, commission_id: int) -> Optional[Tuple]:
        """Получение комиссии по ID с информацией о пользователе"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT ec.*, u.username,
                       (SELECT COUNT(*) FROM commission_experts ce 
                        WHERE ce.commission_id = ec.id) as member_count
                FROM commissions ec
                LEFT JOIN users u ON ec.user_id = u.id
                WHERE ec.id = %s
            """, (commission_id,))
            return cur.fetchone()

    def get_commission_experts_with_roles(self, commission_id: int) -> List[Dict]:
        """Получение экспертов комиссии с ролями (в виде словарей)"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT 
                    e.id as expert_id,
                    e.surname,
                    e.name,
                    e.patronymic,
                    e.position,
                    cr.id as role_id,
                    cr.name as role_name,
                    ce.commission_id
                FROM commission_experts ce
                JOIN expert e ON ce.expert_id = e.id
                JOIN commission_roles cr ON ce.role_id = cr.id
                WHERE ce.commission_id = %s
                ORDER BY 
                    CASE cr.name
                        WHEN 'председатель' THEN 1
                        WHEN 'эксперт' THEN 2
                        ELSE 3
                    END,
                    e.surname, e.name, e.patronymic
            """, (commission_id,))

            experts = []
            for row in cur.fetchall():
                experts.append({
                    'expert_id': row[0],
                    'surname': row[1],
                    'name': row[2],
                    'patronymic': row[3],
                    'position': row[4],
                    'role_id': row[5],
                    'role_name': row[6],
                    'commission_id': row[7]
                })
            return experts

    def get_available_experts(self, commission_id: int, user_id: int) -> List[Tuple]:
        """Получение доступных экспертов (не входящих в комиссию)"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT e.id, e.surname, e.name, e.patronymic, e.position
                FROM expert e
                WHERE e.user_id = %s 
                AND e.id NOT IN (
                    SELECT expert_id FROM commission_experts 
                    WHERE commission_id = %s
                )
                ORDER BY e.surname, e.name, e.patronymic
            """, (user_id, commission_id))
            return cur.fetchall()

    def create_expert_with_return(self, surname: str, name: str, patronymic: str, position: str, user_id: int) -> int:
        """Создание нового эксперта с возвратом ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO expert (surname, name, patronymic, position, user_id, created_date)
                VALUES (%s, %s, %s, %s, %s, %s) RETURNING id
            """, (surname, name, patronymic, position, user_id, datetime.now()))
            return cur.fetchone()[0]

    def add_expert_to_commission_with_check(self, commission_id: int, expert_id: int, role_id: int) -> bool:
        """Добавление эксперта в комиссию с проверкой"""
        try:
            with self.db.get_cursor() as cur:
                # Проверяем, не состоит ли уже эксперт в комиссии
                cur.execute("""
                    SELECT COUNT(*) FROM commission_experts 
                    WHERE commission_id = %s AND expert_id = %s
                """, (commission_id, expert_id))

                if cur.fetchone()[0] > 0:
                    return False

                cur.execute("""
                    INSERT INTO commission_experts (commission_id, expert_id, role_id)
                    VALUES (%s, %s, %s)
                """, (commission_id, expert_id, role_id))
                return True
        except Exception as e:
            print(f"Error adding expert to commission: {e}")
            return False

    def update_commission(self, commission_id: int, name: str, description: str) -> bool:
        """Обновление комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE commissions 
                SET name = %s, description = %s
                WHERE id = %s
            """, (name, description, commission_id))
            return cur.rowcount > 0

    def get_commission_roles_list(self) -> List[Dict]:
        """Получение всех ролей в виде списка словарей"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, name FROM commission_roles ORDER BY id")
            roles = []
            for row in cur.fetchall():
                roles.append({
                    'id': row[0],
                    'name': row[1]
                })
            return roles

    def has_chairman_in_commission(self, commission_id: int) -> bool:
        """Проверка наличия председателя в комиссии"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) 
                FROM commission_experts ce
                JOIN commission_roles cr ON ce.role_id = cr.id
                WHERE ce.commission_id = %s 
                AND LOWER(cr.name) = 'председатель'
            """, (commission_id,))
            result = cur.fetchone()
            return result[0] > 0 if result else False

    def get_experts_by_ids(self, expert_ids: List[int]) -> List[Tuple]:
        """Получение экспертов по списку ID"""
        if not expert_ids:
            return []

        placeholders = ','.join(['%s'] * len(expert_ids))
        with self.db.get_cursor() as cur:
            cur.execute(f"""
                SELECT id, surname, name, patronymic, position
                FROM expert
                WHERE id IN ({placeholders})
                ORDER BY surname, name
            """, expert_ids)
            return cur.fetchall()
