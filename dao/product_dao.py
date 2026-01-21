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

    def create(self, product_data: dict) -> int:
        """Создание нового продукта"""
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
        """Обновление продукта"""
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

    # В классе ProductDAO добавьте метод:
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
