# dao/setting_dao.py
from typing import List, Optional, Dict, Any
from datetime import datetime


class SettingDAO:
    def __init__(self, db):
        self.db = db

    def get_current(self) -> Optional[tuple]:
        """Получение текущих настроек (предполагаем одну запись)"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT s.*, 
                       a.full_name as assignee_name,
                       a.post as assignee_post,  -- Добавлено
                       poa.number as poa_number,
                       poa.date_of_receipt as poa_date,
                       sm.value as sum_value,
                       p.full_name as provost_name,
                       p.post as provost_post    -- Добавлено
                FROM setting s
                LEFT JOIN assignee a ON s.assignee_id = a.id
                LEFT JOIN power_of_attorney poa ON s.power_of_attorney_id = poa.id
                LEFT JOIN sum sm ON s.sum_id = sm.id
                LEFT JOIN provost p ON s.provost_id = p.id  
                ORDER BY s.id DESC
                LIMIT 1
            """)
            return cur.fetchone()

    def save(self, setting_data: dict) -> int:
        """Сохранение настроек (обновление существующей или создание новой записи)"""
        with self.db.get_cursor() as cur:
            # Проверяем, есть ли уже запись
            cur.execute("SELECT id FROM setting LIMIT 1")
            existing = cur.fetchone()

            if existing:
                # Обновляем существующую запись
                cur.execute("""
                    UPDATE setting 
                    SET assignee_id = %s,
                        power_of_attorney_id = %s,
                        sum_id = %s,
                        provost_id = %s
                    WHERE id = %s
                """, (
                    setting_data.get('assignee_id'),
                    setting_data.get('power_of_attorney_id'),
                    setting_data.get('sum_id'),
                    setting_data.get('provost_id'),
                    existing[0]
                ))
                return existing[0]
            else:
                # Создаем новую запись
                cur.execute("""
                    INSERT INTO setting (
                        assignee_id,
                        power_of_attorney_id,
                        sum_id,
                        provost_id
                    ) VALUES (%s, %s, %s, %s)
                    RETURNING id
                """, (
                    setting_data.get('assignee_id'),
                    setting_data.get('power_of_attorney_id'),
                    setting_data.get('sum_id'),
                    setting_data.get('provost_id')
                ))
                return cur.fetchone()[0]

    def update_single(self, field_name: str, value_id: int) -> bool:
        """Обновление одного поля настроек"""
        with self.db.get_cursor() as cur:
            # Проверяем корректность имени поля
            valid_fields = ['assignee_id', 'power_of_attorney_id', 'sum_id', 'provost_id']
            if field_name not in valid_fields:
                return False

            # Получаем текущие настройки
            cur.execute("SELECT id FROM setting LIMIT 1")
            existing = cur.fetchone()

            if existing:
                # Обновляем существующую запись
                cur.execute(f"""
                    UPDATE setting 
                    SET {field_name} = %s
                    WHERE id = %s
                """, (value_id, existing[0]))
            else:
                # Создаем новую запись
                fields = ['assignee_id', 'power_of_attorney_id', 'sum_id', 'provost_id']
                values = [None, None, None, None]

                # Устанавливаем значение в нужное поле
                idx = fields.index(field_name)
                values[idx] = value_id

                cur.execute("""
                    INSERT INTO setting (
                        assignee_id,
                        power_of_attorney_id,
                        sum_id,
                        provost_id
                    ) VALUES (%s, %s, %s, %s)
                """, tuple(values))

            return cur.rowcount > 0

    def delete(self) -> bool:
        """Удаление всех настроек"""
        with self.db.get_cursor() as cur:
            cur.execute("DELETE FROM setting")
            return cur.rowcount > 0

    def get_assignees(self) -> List[tuple]:
        """Получение всех доверенных лиц"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, full_name, post FROM assignee ORDER BY full_name")
            return cur.fetchall()

    def get_powers_of_attorney(self) -> List[tuple]:
        """Получение всех доверенностей"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT id, number, date_of_receipt 
                FROM power_of_attorney 
                ORDER BY date_of_receipt DESC
            """)
            return cur.fetchall()

    def get_sums(self) -> List[tuple]:
        """Получение всех сумм"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, value FROM sum ORDER BY value")
            return cur.fetchall()

    def get_provosts(self) -> List[tuple]:
        """Получение всех проректоров"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, full_name, post FROM provost ORDER BY full_name")
            return cur.fetchall()

    def create_assignee(self, full_name: str, post: str = None) -> int:
        """Создание нового доверенного лица"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO assignee (full_name, post) 
                VALUES (%s, %s)
                RETURNING id
            """, (full_name, post))
            return cur.fetchone()[0]

    def create_power_of_attorney(self, number: str, date_of_receipt: str) -> int:
        """Создание новой доверенности"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO power_of_attorney (number, date_of_receipt) 
                VALUES (%s, %s)
                RETURNING id
            """, (number, date_of_receipt))
            return cur.fetchone()[0]

    def create_sum(self, value: float) -> int:
        """Создание новой суммы"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO sum (value) 
                VALUES (%s)
                RETURNING id
            """, (value,))
            return cur.fetchone()[0]

    def create_provost(self, full_name: str, post: str = None) -> int:
        """Создание нового проректора"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO provost (full_name, post) 
                VALUES (%s, %s)
                RETURNING id
            """, (full_name, post))
            return cur.fetchone()[0]

    def get_setting_by_id(self, setting_id: int) -> Optional[tuple]:
        """Получение настройки по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT s.*, 
                       a.full_name as assignee_name,
                       a.post as assignee_post,  -- Добавлено
                       poa.number as poa_number,
                       poa.date_of_receipt as poa_date,
                       sm.value as sum_value,
                       pr.full_name as provost_name,
                       pr.post as provost_post    -- Добавлено
                FROM setting s
                LEFT JOIN assignee a ON s.assignee_id = a.id
                LEFT JOIN power_of_attorney poa ON s.power_of_attorney_id = poa.id
                LEFT JOIN sum sm ON s.sum_id = sm.id
                LEFT JOIN provost pr ON s.provost_id = pr.id
                WHERE s.id = %s
            """, (setting_id,))
            return cur.fetchone()

    def get_reference_options(self, reference_type: str) -> List[Dict[str, Any]]:
        """Получение опций для выпадающего списка"""
        options = []

        # Добавляем пустой вариант
        options.append({'id': None, 'name': 'Не выбрано', 'display_name': 'Не выбрано'})

        if reference_type == 'assignee':
            rows = self.get_assignees()
            for row in rows:
                display_name = f"{row[1]}" + (f" ({row[2]})" if row[2] else "")
                options.append({
                    'id': row[0],
                    'name': row[1],
                    'display_name': display_name,
                    'post': row[2]
                })
        elif reference_type == 'power_of_attorney':
            rows = self.get_powers_of_attorney()
            for row in rows:
                date_str = row[2].strftime('%d.%m.%Y') if row[2] else ''
                options.append({
                    'id': row[0],
                    'name': row[1],
                    'display_name': f"{row[1]} ({date_str})",
                    'date': row[2]
                })
        elif reference_type == 'sum':
            rows = self.get_sums()
            for row in rows:
                options.append({'id': row[0], 'name': str(row[1]), 'display_name': str(row[1])})
        elif reference_type == 'provost':
            rows = self.get_provosts()
            for row in rows:
                display_name = f"{row[1]}" + (f" ({row[2]})" if row[2] else "")
                options.append({
                    'id': row[0],
                    'name': row[1],
                    'display_name': display_name,
                    'post': row[2]
                })

        return options

    def check_exists(self, table_name: str, field_name: str, value: Any) -> bool:
        """Проверка существования записи в таблице"""
        with self.db.get_cursor() as cur:
            cur.execute(f"SELECT id FROM {table_name} WHERE {field_name} = %s", (value,))
            return cur.fetchone() is not None
