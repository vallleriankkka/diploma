# database.py
import psycopg2
from contextlib import contextmanager
from typing import Generator, Optional, Dict, Any
from config import load_database_config


class Database:
    """Класс для работы с базой данных PostgreSQL"""

    # Реализация синглтона
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(Database, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if not self._initialized:
            self._config = None
            self._initialized = True

    def set_config(self, config: Dict[str, Any]) -> None:
        """Установка конфигурации базы данных (должна быть вызвана перед использованием)"""
        self._config = config

    def _get_config(self) -> Dict[str, Any]:
        """Получение конфигурации базы данных"""
        if self._config is None:
            # Загружаем конфигурацию из config.py
            self._config = load_database_config()
        return self._config

    def get_connection(self) -> psycopg2.extensions.connection:
        """Создает подключение к базе данных"""
        config = self._get_config()
        return psycopg2.connect(**config)

    @contextmanager
    def get_cursor(self, connection: Optional[psycopg2.extensions.connection] = None) -> Generator[
        psycopg2.extensions.cursor, None, None]:
        """Контекстный менеджер для работы с курсором"""
        conn = connection or self.get_connection()
        cur = conn.cursor()
        try:
            yield cur
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            cur.close()
            if not connection:
                conn.close()

    @classmethod
    def get_instance(cls) -> 'Database':
        """Получение экземпляра базы данных"""
        if cls._instance is None:
            cls._instance = Database()
        return cls._instance


# Фабричная функция для удобного создания Database с конфигурацией
def create_database(config: Optional[Dict[str, Any]] = None) -> Database:
    """
    Создание и настройка экземпляра Database
    """
    db = Database.get_instance()
    if config:
        db.set_config(config)
    return db

