import os
import io
from datetime import datetime
from docx import Document
from flask import current_app, send_file


class AbstractGenerator:
    """Генератор документа Аннотация"""

    def __init__(self, template_name='abstract_template.docx'):
        self.template_name = template_name
        self.template_path = self._get_template_path()

    def _get_template_path(self):
        """Получение пути к шаблону"""
        return os.path.join(
            current_app.root_path,
            'templates',
            'documents',
            self.template_name
        )

    def generate(self, product_data, authors_data, settings_data=None):
        """Генерация документа Аннотация"""
        try:
            if not os.path.exists(self.template_path):
                print(f"Шаблон не найден: {self.template_path}")
                return None

            doc = Document(self.template_path)
            self._fill_template(doc, product_data, authors_data, settings_data)
            return doc

        except Exception as e:
            print(f"Ошибка при генерации аннотации: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _fill_template(self, doc, product_data, authors_data, settings_data=None):
        """Заполнение шаблона данными"""
        print("\n" + "=" * 50)
        print("DEBUG: START FILL ABSTRACT TEMPLATE")
        print("=" * 50)

        # Извлекаем данные продукта
        product_info = self._extract_product_data(product_data)
        print(f"DEBUG: Product data extracted: {product_info}")

        # Извлекаем данные авторов
        authors_info = self._extract_authors_data(authors_data)
        print(f"DEBUG: Authors data extracted: {authors_info}")

        # Словарь замен для аннотации
        replacements = {
            # Название продукта
            'PRODUCT_NAME': product_info.get('name', 'Не указано'),

            # Аннотация/описание
            'ANNOTATION': product_info.get('description', 'Не указано'),

            'DESCRIPTION': product_info.get('description', 'Не указано'),

            # Тип ЭВМ
            'COMPUTER_TYPE': product_info.get('computer_type', 'Не указан'),

            # Язык программирования
            'PROGRAMMING_LANGUAGE': product_info.get('programming_language', 'Не указан'),

            'LANGUAGE': product_info.get('programming_language', 'Не указан'),

            # Операционная система
            'OPERATING_SYSTEM': product_info.get('operating_system', 'Не указана'),

            'OS': product_info.get('operating_system', 'Не указана'),

            # Объём программы
            'VOLUME': product_info.get('volume', 'Не указан'),

            'PROGRAM_VOLUME': product_info.get('volume', 'Не указан'),

            # Авторы
            'AUTHORS': authors_info.get('full_names', 'Авторы не указаны'),

            'AUTHOR_FULL_NAME': authors_info.get('full_names', 'Авторы не указаны'),

            # Должности/роли авторов
            'AUTHORS_ROLES': authors_info.get('roles', 'Роли не указаны'),

        }

        print(f"\nDEBUG: Key replacements for abstract:")
        print(f"  PRODUCT_NAME: '{replacements['PRODUCT_NAME']}'")
        print(f"  ANNOTATION: '{replacements['ANNOTATION'][:50]}...'")
        print(f"  COMPUTER_TYPE: '{replacements['COMPUTER_TYPE']}'")
        print(f"  LANGUAGE: '{replacements['LANGUAGE']}'")
        print(f"  OS: '{replacements['OS']}'")
        print(f"  VOLUME: '{replacements['VOLUME']}'")
        print(f"  AUTHORS: '{replacements['AUTHORS']}'")

        print("\n" + "=" * 50)
        print("DEBUG: END FILL ABSTRACT TEMPLATE")
        print("=" * 50)

        self._replace_in_paragraphs_preserve_formatting(doc, replacements)
        self._replace_in_tables_preserve_formatting(doc, replacements)

    def _extract_product_data(self, product_data):
        """Извлечение данных продукта из DAO результата"""
        if not product_data or not isinstance(product_data, tuple):
            return {
                'name': '',
                'description': '',
                'computer_type': '',
                'programming_language': '',
                'operating_system': '',
                'volume': ''
            }

        print(f"\nDEBUG: Product data structure:")
        for i, item in enumerate(product_data):
            print(f"  [{i}]: {repr(item)}")

        result = {
            'name': '',
            'description': '',
            'computer_type': '',
            'programming_language': '',
            'operating_system': '',
            'volume': ''
        }

        if len(product_data) > 1 and product_data[1]:
            result['name'] = str(product_data[1]).strip()
            print(f"DEBUG: Name from index 1: '{result['name']}'")

        if len(product_data) > 2 and product_data[2]:
            result['description'] = str(product_data[2]).strip()
            print(f"DEBUG: Description from index 2: '{result['description']}'")

        # Компьютерный тип - индекс 9
        if len(product_data) > 9 and product_data[9]:
            result['computer_type'] = str(product_data[9]).strip()
            print(f"DEBUG: Computer type from index 9: '{result['computer_type']}'")

        # Язык программирования - индекс 10
        if len(product_data) > 10 and product_data[10]:
            result['programming_language'] = str(product_data[10]).strip()
            print(f"DEBUG: Programming language from index 10: '{result['programming_language']}'")

        # Операционная система - индекс 11
        if len(product_data) > 11 and product_data[11]:
            result['operating_system'] = str(product_data[11]).strip()
            print(f"DEBUG: Operating system from index 11: '{result['operating_system']}'")

        # Объем - индекс 7
        if len(product_data) > 3 and product_data[3]:
            volume = product_data[3]
            if isinstance(volume, (int, float)):
                result['volume'] = str(volume)
            elif isinstance(volume, str):
                result['volume'] = volume.strip()
            else:
                result['volume'] = str(volume) if volume else ''
            print(f"DEBUG: Volume from index 7: '{result['volume']}'")

        return result

    def _extract_authors_data(self, authors_data):
        """Извлечение данных авторов для аннотации"""
        if not authors_data:
            return {'full_names': 'Авторы не указаны'}

        # Просто берем ФИО из первых 5 авторов
        full_names = []
        for i, author in enumerate(authors_data[:5]):  # Ограничим 5 авторами
            if isinstance(author, tuple) and len(author) > 9:
                # Пробуем разные комбинации
                parts = []
                for idx in [7, 8, 9]:  # Разные возможные индексы
                    if idx < len(author) and author[idx] and isinstance(author[idx], str):
                        part = str(author[idx]).strip()
                        if part and part not in parts:
                            parts.append(part)

                if parts:
                    full_name = ' '.join(parts)
                    full_names.append(full_name)

        if full_names:
            return {'full_names': ', '.join(full_names)}
        else:
            return {'full_names': 'Авторы не указаны'}

    def _get_month_name_ru(self, month_num):
        """Название месяца на русском в родительном падеже"""
        try:
            months = {
                1: 'января',
                2: 'февраля',
                3: 'марта',
                4: 'апреля',
                5: 'мая',
                6: 'июня',
                7: 'июля',
                8: 'августа',
                9: 'сентября',
                10: 'октября',
                11: 'ноября',
                12: 'декабря'
            }
            return months.get(int(month_num), '')
        except:
            return ''

    def _replace_in_paragraphs_preserve_formatting(self, doc, replacements):
        """Замена в параграфах"""
        for paragraph in doc.paragraphs:
            if not paragraph.text:
                continue

            for old_text, new_text in replacements.items():
                if old_text in paragraph.text:
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

    def _replace_in_tables_preserve_formatting(self, doc, replacements):
        """Замена в таблицах"""
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if not paragraph.text:
                            continue

                        for old_text, new_text in replacements.items():
                            if old_text in paragraph.text:
                                for run in paragraph.runs:
                                    if old_text in run.text:
                                        run.text = run.text.replace(old_text, new_text)
