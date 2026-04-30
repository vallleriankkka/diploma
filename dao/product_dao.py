from typing import List, Optional


class ProductDAO:
    def __init__(self, db):
        self.db = db

    def get_all(self, user_id: Optional[int] = None) -> List[tuple]:
        """Получение всех продуктов"""
        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                    SELECT p.*, ct.name as computer_type, pl.name as programming_language, 
                           os.name as operating_system
                    FROM product p 
                    LEFT JOIN computer_types ct ON p.computer_type_id = ct.id
                    LEFT JOIN programming_languages pl ON p.programming_language_id = pl.id
                    LEFT JOIN os_types os ON p.operating_system_id = os.id
                    WHERE p.user_id = %s
                    ORDER BY p.name
                """, (user_id,))
            else:
                cur.execute("""
                    SELECT p.*, ct.name as computer_type, pl.name as programming_language, 
                           os.name as operating_system, u.username
                    FROM product p 
                    LEFT JOIN computer_types ct ON p.computer_type_id = ct.id
                    LEFT JOIN programming_languages pl ON p.programming_language_id = pl.id
                    LEFT JOIN os_types os ON p.operating_system_id = os.id
                    LEFT JOIN users u ON p.user_id = u.id
                    ORDER BY p.name
                """)
            return cur.fetchall()

    def get_by_id(self, product_id: int) -> Optional[tuple]:
        """Получение продукта по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT p.*, ct.name as computer_type, pl.name as programming_language, 
                       os.name as operating_system
                FROM product p 
                LEFT JOIN computer_types ct ON p.computer_type_id = ct.id
                LEFT JOIN programming_languages pl ON p.programming_language_id = pl.id
                LEFT JOIN os_types os ON p.operating_system_id = os.id
                WHERE p.id = %s
            """, (product_id,))
            return cur.fetchone()

    def check_compatibility(self, os_id: int, computer_type_id: int) -> bool:
        """Проверка совместимости ОС и типа ЭВМ"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT COUNT(*) 
                FROM os_computer_type_compatibility 
                WHERE os_type_id = %s AND computer_type_id = %s
            """, (os_id, computer_type_id))
            count = cur.fetchone()[0]
            return count > 0

    def get_compatible_computer_types(self, os_id: int) -> List[dict]:
        """Получение всех типов ЭВМ, совместимых с данной ОС"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT ct.id, ct.name
                FROM computer_types ct
                JOIN os_computer_type_compatibility octc ON ct.id = octc.computer_type_id
                WHERE octc.os_type_id = %s
                ORDER BY ct.name
            """, (os_id,))
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]

    def get_compatible_os_types(self, computer_type_id: int) -> List[dict]:
        """Получение всех ОС, совместимых с данным типом ЭВМ"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT ot.id, ot.name
                FROM os_types ot
                JOIN os_computer_type_compatibility octc ON ot.id = octc.os_type_id
                WHERE octc.computer_type_id = %s
                ORDER BY ot.name
            """, (computer_type_id,))
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]

    def add_compatibility(self, os_id: int, computer_type_id: int) -> bool:
        """Добавление совместимости между ОС и типом ЭВМ"""
        try:
            with self.db.get_cursor() as cur:
                cur.execute("""
                    INSERT INTO os_computer_type_compatibility (os_type_id, computer_type_id)
                    VALUES (%s, %s)
                    ON CONFLICT (os_type_id, computer_type_id) DO NOTHING
                """, (os_id, computer_type_id))
                return cur.rowcount > 0
        except Exception:
            return False

    def remove_compatibility(self, os_id: int, computer_type_id: int) -> bool:
        """Удаление совместимости между ОС и типом ЭВМ"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                DELETE FROM os_computer_type_compatibility 
                WHERE os_type_id = %s AND computer_type_id = %s
            """, (os_id, computer_type_id))
            return cur.rowcount > 0

    def get_all_compatibilities(self) -> List[dict]:
        """Получение всех связей совместимости"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT 
                    ot.id as os_id,
                    ot.name as os_name,
                    ct.id as computer_type_id,
                    ct.name as computer_type_name
                FROM os_computer_type_compatibility octc
                JOIN os_types ot ON octc.os_type_id = ot.id
                JOIN computer_types ct ON octc.computer_type_id = ct.id
                ORDER BY ot.name, ct.name
            """)
            return [
                {
                    'os_id': row[0],
                    'os_name': row[1],
                    'computer_type_id': row[2],
                    'computer_type_name': row[3]
                }
                for row in cur.fetchall()
            ]

    def create(self, product_data: dict) -> int:
        """Создание нового продукта с проверкой совместимости"""
        # Проверяем совместимость, если указаны оба параметра
        if product_data.get('operating_system_id') and product_data.get('computer_type_id'):
            if not self.check_compatibility(
                    product_data['operating_system_id'],
                    product_data['computer_type_id']
            ):
                raise ValueError("Операционная система не совместима с выбранным типом ЭВМ")

        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO product (name, description, computer_type_id, programming_language_id, 
                operating_system_id, volume, user_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (
                product_data['name'],
                product_data['description'],
                product_data['computer_type_id'],
                product_data['programming_language_id'],
                product_data['operating_system_id'],
                product_data['volume'],
                product_data['user_id']
            ))
            return cur.fetchone()[0]

    def update(self, product_id: int, product_data: dict) -> bool:
        """Обновление продукта с проверкой совместимости"""
        # Проверяем совместимость, если указаны оба параметра
        if product_data.get('operating_system_id') and product_data.get('computer_type_id'):
            if not self.check_compatibility(
                    product_data['operating_system_id'],
                    product_data['computer_type_id']
            ):
                raise ValueError("Операционная система не совместима с выбранным типом ЭВМ")

        with self.db.get_cursor() as cur:
            cur.execute("""
                UPDATE product SET name=%s, description=%s, computer_type_id=%s, 
                programming_language_id=%s, operating_system_id=%s, volume=%s
                WHERE id=%s
            """, (
                product_data['name'],
                product_data['description'],
                product_data['computer_type_id'],
                product_data['programming_language_id'],
                product_data['operating_system_id'],
                product_data['volume'],
                product_id
            ))
            return cur.rowcount > 0

    def delete(self, product_id: int) -> bool:
        """Удаление продукта"""
        with self.db.get_cursor() as cur:
            cur.execute("DELETE FROM product WHERE id = %s", (product_id,))
            return cur.rowcount > 0

    def get_by_user(self, product_id: int, user_id: int) -> Optional[int]:
        """Проверка принадлежности продукта пользователю"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT user_id FROM product WHERE id = %s", (product_id,))
            result = cur.fetchone()
            return result[0] if result else None

    def get_product_authors(self, product_id):
        """Получение авторов продукта"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT 
                    au.id,
                    au.surname,
                    au.name,
                    au.patronymic,
                    er.name as role_name,
                    aa.participation_percent,
                    aa.time_spent,
                    aa.description
                FROM product p
                JOIN applications a ON p.id = a.product_id
                JOIN applicationauthor aa ON a.id = aa.application_id
                JOIN authors au ON aa.author_id = au.id
                LEFT JOIN expert_roles er ON aa.role_id = er.id
                WHERE p.id = %s
                ORDER BY au.surname, au.name
            """, (product_id,))
            return cur.fetchall()

    def get_all_programming_languages(self) -> List[dict]:
        """Получение всех языков программирования"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                    SELECT id, name 
                    FROM programming_languages 
                    ORDER BY name
                """)
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]

    def get_all_os_types(self) -> List[dict]:
        """Получение всех типов ОС"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                    SELECT id, name 
                    FROM os_types 
                    ORDER BY name
                """)
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]

    def get_all_computer_types(self) -> List[dict]:
        """Получение всех типов ЭВМ"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                    SELECT id, name 
                    FROM computer_types 
                    ORDER BY name
                """)
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]

    def create_reference(self, reference_type: str, name: str) -> dict:
        """Создание новой записи в справочнике"""
        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            raise ValueError(f"Неизвестный тип справочника: {reference_type}")

        with self.db.get_cursor() as cur:
            cur.execute(f"""
                    INSERT INTO {table_name} (name) 
                    VALUES (%s) 
                    RETURNING id, name
                """, (name,))
            result = cur.fetchone()
            return {'id': result[0], 'name': result[1]}

    def update_reference(self, reference_type: str, ref_id: int, name: str) -> bool:
        """Обновление записи в справочнике"""
        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            raise ValueError(f"Неизвестный тип справочника: {reference_type}")

        with self.db.get_cursor() as cur:
            cur.execute(f"""
                    UPDATE {table_name} 
                    SET name = %s 
                    WHERE id = %s
                """, (name, ref_id))
            return cur.rowcount > 0

    def delete_reference(self, reference_type: str, ref_id: int) -> bool:
        """Удаление записи из справочника"""
        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            raise ValueError(f"Неизвестный тип справочника: {reference_type}")

        # Проверяем, используется ли запись в продуктах
        product_table_mapping = {
            'programming_language': 'programming_language_id',
            'os_type': 'operating_system_id',
            'computer_type': 'computer_type_id'
        }

        product_field = product_table_mapping.get(reference_type)
        if product_field:
            with self.db.get_cursor() as cur:
                cur.execute(f"""
                        SELECT COUNT(*) FROM product 
                        WHERE {product_field} = %s
                    """, (ref_id,))
                usage_count = cur.fetchone()[0]

                if usage_count > 0:
                    raise ValueError(f"Этот элемент используется в {usage_count} продуктах и не может быть удален")

        with self.db.get_cursor() as cur:
            cur.execute(f"""
                    DELETE FROM {table_name} 
                    WHERE id = %s
                """, (ref_id,))
            return cur.rowcount > 0