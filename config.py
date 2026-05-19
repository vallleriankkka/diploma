import os
from typing import Dict, Any
from pathlib import Path


def load_database_config() -> Dict[str, Any]:
    """
    Загрузка конфигурации базы данных.
    Приоритет: переменные окружения > файл config.json > значения по умолчанию
    """
    # Значения по умолчанию
    default_config = {
        'host': 'localhost',
        'port': 5432,
        'database': 'diplom',
        'user': 'postgres',
        'password': '0109'
    }



    # Переменные окружения
    env_config = {
        'host': os.getenv('DB_HOST'),
        'port': os.getenv('DB_PORT'),
        'database': os.getenv('DB_NAME'),
        'user': os.getenv('DB_USER'),
        'password': os.getenv('DB_PASSWORD')
    }

    # Обновляем конфигурацию значениями из переменных окружения (если они есть)
    for key, value in env_config.items():
        if value is not None:
            if key == 'port':
                default_config[key] = int(value)
            else:
                default_config[key] = value

    return default_config