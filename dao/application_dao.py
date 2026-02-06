from typing import List, Optional, Tuple, Dict


class ApplicationDAO:
    def __init__(self, db):
        self.db = db
        self._status_cache = None  # Кэш статусов

    def _get_statuses(self) -> Dict[str, int]:
        """Получение статусов из базы данных"""
        if self._status_cache is None:
            with self.db.get_cursor() as cur:
                cur.execute("SELECT id, name FROM application_statuses ORDER BY id")
                statuses = cur.fetchall()
                self._status_cache = {
                    "Новый": statuses[0][0] if len(statuses) > 0 else 1,
                    "Отправлен на рассмотрение": statuses[1][0] if len(statuses) > 1 else 2,
                    "Подготовлен": statuses[2][0] if len(statuses) > 2 else 3,
                    "Отправлен на доработку": statuses[3][0] if len(statuses) > 3 else 4,
                    "Принят": statuses[4][0] if len(statuses) > 4 else 5
                }
        return self._status_cache

    @property
    def STATUS_NEW(self) -> int:
        return self._get_statuses()["Новый"]

    @property
    def STATUS_FOR_REVIEW(self) -> int:
        return self._get_statuses()["Отправлен на рассмотрение"]

    @property
    def STATUS_PREPARED(self) -> int:
        return self._get_statuses()["Подготовлен"]

    @property
    def STATUS_FOR_REVISION(self) -> int:
        return self._get_statuses()["Отправлен на доработку"]

    @property
    def STATUS_ACCEPTED(self) -> int:
        return self._get_statuses()["Принят"]

    def get_all(self, user_id: Optional[int] = None) -> List[tuple]:
        """Получение всех заявок"""
        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                    SELECT a.*, p.name as product_name, s.name as status_name 
                    FROM application a 
                    LEFT JOIN product p ON a.product_id = p.id 
                    LEFT JOIN application_statuses s ON a.status_id = s.id
                    WHERE a.user_id = %s
                    ORDER BY a.created_date DESC
                """, (user_id,))
            else:
                cur.execute("""
                    SELECT a.*, p.name as product_name, u.username, s.name as status_name
                    FROM application a 
                    LEFT JOIN product p ON a.product_id = p.id 
                    LEFT JOIN users u ON a.user_id = u.id
                    LEFT JOIN application_statuses s ON a.status_id = s.id
                    ORDER BY a.created_date DESC
                """)
            return cur.fetchall()

    def get_by_id(self, application_id: int, user_id: Optional[int] = None) -> Optional[tuple]:
        """Получение заявки по ID"""
        with self.db.get_cursor() as cur:
            if user_id:
                cur.execute("""
                      SELECT a.*, p.name as product_name,
                             ec.name as commission_name, s.name as status_name,
                             s.id as status_id, u.username
                      FROM application a 
                      LEFT JOIN product p ON a.product_id = p.id 
                      LEFT JOIN commissions ec ON a.commission_id = ec.id
                      LEFT JOIN application_statuses s ON a.status_id = s.id
                      LEFT JOIN users u ON a.user_id = u.id
                      WHERE a.id = %s AND a.user_id = %s
                  """, (application_id, user_id))
            else:
                cur.execute("""
                      SELECT a.*, p.name as product_name, u.username,
                             ec.name as commission_name, s.name as status_name,
                             s.id as status_id
                      FROM application a 
                      LEFT JOIN product p ON a.product_id = p.id 
                      LEFT JOIN users u ON a.user_id = u.id
                      LEFT JOIN commissions ec ON a.commission_id = ec.id
                      LEFT JOIN application_statuses s ON a.status_id = s.id
                      WHERE a.id = %s
                  """, (application_id,))
            return cur.fetchone()

    def create(self, application_data: dict) -> int:
        """Создание новой заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO application (title, product_id, user_id, commission_id, status_id, nir) 
                VALUES (%s, %s, %s, %s, %s, %s) RETURNING id
            """, (
                application_data['title'],
                application_data['product_id'],
                application_data['user_id'],
                application_data.get('commission_id'),
                self.STATUS_NEW,  # Используем динамический статус
                application_data['nir']
            ))
            return cur.fetchone()[0]

    def update_status(self, application_id: int, status_id: int) -> bool:
        """Обновление статуса заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("UPDATE application SET status_id = %s WHERE id = %s",
                        (status_id, application_id))
            return cur.rowcount > 0

    def delete(self, application_id: int) -> bool:
        """Удаление заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("DELETE FROM application WHERE id = %s", (application_id,))
            return cur.rowcount > 0

    def get_by_user(self, application_id: int, user_id: int) -> Optional[int]:
        """Получение пользователя заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT user_id FROM application WHERE id = %s", (application_id,))
            result = cur.fetchone()
            return result[0] if result else None

    def get_application_authors(self, application_id: int) -> List[tuple]:
        """Получение авторов заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT aa.*, 
                       au.surname, au.name, au.patronymic,
                       er.name as role_name
                FROM applicationauthor aa
                LEFT JOIN author au ON aa.author_id = au.id
                LEFT JOIN expert_roles er ON aa.role_id = er.id
                WHERE aa.application_id = %s
            """, (application_id,))
            return cur.fetchall()

    def get_application_authors_detailed(self, application_id):
        """Получение авторов заявки с полными данными (с образованием и должностью)"""
        with self.db.get_cursor() as cursor:
            cursor.execute("""
                    SELECT 
                        aa.id,
                        aa.application_id,
                        aa.author_id,
                        aa.participation_percent,
                        aa.time_spent,
                        aa.description,
                        aa.equipment,
                        au.surname,
                        au.name,
                        au.patronymic,
                        au.birth_date,
                        au.citizenship,
                        au.address,
                        au.is_student,
                        au.position,
                        au.education_level_id,
                        el.name as education_level_name,
                        er.name as role_name,
                        au.department
                    FROM applicationauthor aa
                    JOIN author au ON aa.author_id = au.id
                    LEFT JOIN education_level el ON au.education_level_id = el.id
                    LEFT JOIN expert_roles er ON aa.role_id = er.id
                    WHERE aa.application_id = %s
                    ORDER BY aa.id
                """, (application_id,))
            return cursor.fetchall()


    def add_author_to_application(self, application_id: int, author_data: dict) -> None:
        """Добавление автора в заявку"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                INSERT INTO applicationauthor (application_id, author_id, role_id, participation_percent, time_spent, description, equipment)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (
                application_id,
                author_data['author_id'],
                author_data['role_id'],
                author_data['participation_percent'],
                author_data.get('time_spent'),
                author_data['description'],
                author_data['equipment']
            ))

    def can_user_edit(self, application_id: int, user_id: int, is_admin: bool = False) -> bool:
        """
        Проверяет, может ли пользователь редактировать заявку.
        """
        with self.db.get_cursor() as cur:
            # Проверяем, принадлежит ли заявка пользователю
            cur.execute("""
                SELECT user_id, status_id 
                FROM application 
                WHERE id = %s
            """, (application_id,))

            result = cur.fetchone()
            if not result:
                return False

            app_user_id, status_id = result

            # Если заявка не принадлежит пользователю
            if app_user_id != user_id:
                return False

            # Админ может редактировать всегда
            if is_admin:
                return True

            # Обычный пользователь может редактировать только при статусах 1 и 4
            return status_id in (self.STATUS_NEW, self.STATUS_FOR_REVISION)

    def can_user_submit_for_review(self, application_id: int, user_id: int) -> bool:
        """
        Проверяет, может ли пользователь отправить заявку на рассмотрение.
        """
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT user_id, status_id 
                FROM application 
                WHERE id = %s
            """, (application_id,))

            result = cur.fetchone()
            if not result:
                return False

            app_user_id, status_id = result

            # Проверяем принадлежность
            if app_user_id != user_id:
                return False

            # Проверяем допустимые статусы для отправки
            return status_id in (self.STATUS_NEW, self.STATUS_FOR_REVISION)

    def submit_for_review(self, application_id: int, user_id: int) -> bool:
        """
        Отправляет заявку на рассмотрение (меняет статус на "Отправлен на рассмотрение").
        """
        with self.db.get_cursor() as cur:
            # Проверяем возможность отправки
            if not self.can_user_submit_for_review(application_id, user_id):
                return False

            # Меняем статус
            cur.execute("""
                UPDATE application 
                SET status_id = %s 
                WHERE id = %s AND user_id = %s
            """, (self.STATUS_FOR_REVIEW, application_id, user_id))

            return cur.rowcount > 0

    def update_application(self, application_id: int, user_id: int,
                           update_data: dict, is_admin: bool = False) -> bool:
        """
        Обновление данных заявки с проверкой прав.
        """
        with self.db.get_cursor() as cur:
            # Проверяем возможность редактирования
            if not self.can_user_edit(application_id, user_id, is_admin):
                return False

            # Формируем SQL запрос для обновления
            fields = []
            values = []

            for key, value in update_data.items():
                fields.append(f"{key} = %s")
                values.append(value)

            if not fields:
                return False

            values.append(application_id)

            # Добавляем проверку пользователя если не админ
            if not is_admin:
                sql_condition = " AND user_id = %s"
                values.append(user_id)
            else:
                sql_condition = ""

            sql = f"""
                UPDATE application 
                SET {', '.join(fields)}
                WHERE id = %s{sql_condition}
            """

            cur.execute(sql, values)
            return cur.rowcount > 0

    def get_status_name(self, status_id: int) -> str:
        """Получение названия статуса по ID"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT name FROM application_statuses WHERE id = %s", (status_id,))
            result = cur.fetchone()
            return result[0] if result else "Неизвестно"

    def get_all_statuses(self) -> List[tuple]:
        """Получение всех статусов из базы"""
        with self.db.get_cursor() as cur:
            cur.execute("SELECT id, name FROM application_statuses ORDER BY id")
            return cur.fetchall()

    def get_editable_statuses_for_user(self) -> List[int]:
        """Получение списка ID статусов, при которых пользователь может редактировать"""
        return [self.STATUS_NEW, self.STATUS_FOR_REVISION]

    def get_statuses_for_submission(self) -> List[int]:
        """Получение списка ID статусов, при которых можно отправить на рассмотрение"""
        return [self.STATUS_NEW, self.STATUS_FOR_REVISION]

    def remove_author_from_application(self, application_id: int, author_id: int) -> bool:
        """Удаление автора из заявки"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                DELETE FROM applicationauthor 
                WHERE application_id = %s AND author_id = %s
            """, (application_id, author_id))
            return cur.rowcount > 0

    def update_application_author(self, application_id: int, author_id: int,
                                  update_data: dict) -> bool:
        """Обновление данных автора в заявке"""
        with self.db.get_cursor() as cur:
            fields = []
            values = []

            for key, value in update_data.items():
                fields.append(f"{key} = %s")
                values.append(value)

            if not fields:
                return False

            values.append(application_id)
            values.append(author_id)

            sql = f"""
                UPDATE applicationauthor 
                SET {', '.join(fields)}
                WHERE application_id = %s AND author_id = %s
            """

            cur.execute(sql, values)
            return cur.rowcount > 0

    def clear_status_cache(self):
        """Очистка кэша статусов (например, после изменения статусов в БД)"""
        self._status_cache = None

    def get_application_by_id(self, application_id):
        """Получить заявку по ID без проверки прав (для генератора документов)"""
        print(f"\n[ApplicationDAO.get_application_by_id] Получение заявки {application_id}")

        try:
            with self.db.get_cursor() as cur:
                cur.execute("""
                    SELECT a.*, p.name as product_name,
                           ec.name as commission_name, s.name as status_name,
                           s.id as status_id
                    FROM application a 
                    LEFT JOIN product p ON a.product_id = p.id 
                    LEFT JOIN commissions ec ON a.commission_id = ec.id
                    LEFT JOIN application_statuses s ON a.status_id = s.id
                    WHERE a.id = %s
                """, (application_id,))

                application = cur.fetchone()

                if application:
                    print(
                        f"[ApplicationDAO.get_application_by_id] Заявка найдена: ID={application[0]}, Название='{application[1]}'")
                else:
                    print(f"[ApplicationDAO.get_application_by_id] Заявка с ID={application_id} не найдена")

                return application

        except Exception as e:
            print(f"[ApplicationDAO.get_application_by_id] Ошибка: {e}")
            import traceback
            traceback.print_exc()
            return None
    def get_by_id_admin(self, application_id):
        """Получить заявку для администратора (без проверки владельца)"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT a.*, p.name as product_name,
                       ec.name as commission_name, s.name as status_name,
                       s.id as status_id
                FROM applications a 
                LEFT JOIN product p ON a.product_id = p.id 
                LEFT JOIN commissions ec ON a.commission_id = ec.id
                LEFT JOIN application_statuses s ON a.status_id = s.id
                WHERE a.id = %s
            """, (application_id,))
            return cur.fetchone()

    def get_all_author_roles(self) -> List[dict]:
        """Получение всех ролей авторов"""
        with self.db.get_cursor() as cur:
            cur.execute("""
                SELECT id, name 
                FROM expert_roles 
                ORDER BY name
            """)
            return [{'id': row[0], 'name': row[1]} for row in cur.fetchall()]
