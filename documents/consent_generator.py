import os
import io
from datetime import datetime, date
from docx import Document
from flask import current_app, send_file


class ConsentGenerator:
    """Генератор документа согласия (по одному документу на каждого автора)"""

    def __init__(self, template_name='consent_template.docx'):
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

    def generate_single_author(self, application, author, settings_data=None, author_index=0):
        """Генерация документа для одного автора"""
        try:
            if not os.path.exists(self.template_path):
                print(f"Шаблон не найден: {self.template_path}")
                return None

            if not author:
                print("Нет данных об авторе")
                return None

            doc = Document(self.template_path)
            self._fill_template_for_single_author(doc, application, author, settings_data, author_index)

            # Сохраняем документ в буфер
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            return doc_buffer

        except Exception as e:
            print(f"Ошибка при генерации согласия: {e}")
            import traceback
            traceback.print_exc()
            return None

    def generate_all_authors(self, application, authors, settings_data=None):
        """Генерация документов для всех авторов (возвращает список буферов)"""
        if not authors:
            return []

        documents = []
        for i, author in enumerate(authors):
            doc_buffer = self.generate_single_author(application, author, settings_data, i)
            if doc_buffer:
                documents.append(doc_buffer)

        return documents

    def _fill_template_for_single_author(self, doc, application, author, settings_data=None, author_index=0):
        """Заполнение шаблона данными для одного автора"""
        print(f"\n{'=' * 50}")
        print(f"DEBUG: FILL TEMPLATE FOR AUTHOR #{author_index}")
        print('=' * 50)

        # 1. Проректор - имя и должность
        provost_name = self._extract_provost_name(settings_data)
        provost_post = self._extract_provost_post(settings_data)  # НОВЫЙ МЕТОД

        print(f"DEBUG: Extracted provost name: '{provost_name}'")
        print(f"DEBUG: Extracted provost post: '{provost_post}'")

        # 2. Данные заявки
        app_data = self._extract_application_data(application)
        print(f"DEBUG: Application data: {app_data}")

        # 3. Данные одного автора
        author_data = self._extract_single_author_data_detailed(author, author_index)
        print(f"DEBUG: Author data: {author_data}")

        # Словарь замен (только для одного автора)
        replacements = {
            # Программа/продукт
            '{{PROGRAM_NAME}}': app_data.get('product_name', 'Не указано'),
            '{PROGRAM_NAME}': app_data.get('product_name', 'Не указано'),
            'PROGRAM_NAME': app_data.get('product_name', 'Не указано'),

            '{{PRODUCT_NAME}}': app_data.get('product_name', 'Не указано'),
            '{PRODUCT_NAME}': app_data.get('product_name', 'Не указано'),
            'PRODUCT_NAME': app_data.get('product_name', 'Не указано'),

            # ФИО одного автора
            '{{AUTHOR_FULL_NAME}}': author_data.get('full_name', 'Автор не указан'),
            '{AUTHOR_FULL_NAME}': author_data.get('full_name', 'Автор не указан'),
            'AUTHOR_FULL_NAME': author_data.get('full_name', 'Автор не указан'),

            '{{AUTHORS_NAMES}}': author_data.get('full_name', 'Автор не указан'),
            '{AUTHORS_NAMES}': author_data.get('full_name', 'Автор не указан'),
            'AUTHORS_NAMES': author_data.get('full_name', 'Автор не указан'),

            # Дата рождения (одна дата)
            '{{DAY}}': author_data.get('birth_day', ''),
            '{DAY}': author_data.get('birth_day', ''),
            'DAY': author_data.get('birth_day', ''),

            '{{DAY_PADDED}}': author_data.get('birth_day_padded', ''),
            '{DAY_PADDED}': author_data.get('birth_day_padded', ''),
            'DAY_PADDED': author_data.get('birth_day_padded', ''),

            '{{MONTH}}': author_data.get('birth_month', ''),
            '{MONTH}': author_data.get('birth_month', ''),
            'MONTH': author_data.get('birth_month', ''),

            '{{MONTH_NUM}}': author_data.get('birth_month_num', ''),
            '{MONTH_NUM}': author_data.get('birth_month_num', ''),
            'MONTH_NUM': author_data.get('birth_month_num', ''),

            '{{YEAR}}': author_data.get('birth_year', ''),
            '{YEAR}': author_data.get('birth_year', ''),
            'YEAR': author_data.get('birth_year', ''),

            # Полная дата рождения
            '{{BIRTH_DATE}}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            '{BIRTH_DATE}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            'BIRTH_DATE': author_data.get('birth_date_full', 'Дата рождения не указана'),

            '{{BIRTH_DATES}}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            '{BIRTH_DATES}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            'BIRTH_DATES': author_data.get('birth_date_full', 'Дата рождения не указана'),

            '{{AUTHORS_BIRTH_DATES}}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            '{AUTHORS_BIRTH_DATES}': author_data.get('birth_date_full', 'Дата рождения не указана'),
            'AUTHORS_BIRTH_DATES': author_data.get('birth_date_full', 'Дата рождения не указана'),

            # Гражданство (одно)
            '{{CITIZENSHIP}}': author_data.get('citizenship', 'Российская Федерация'),
            '{CITIZENSHIP}': author_data.get('citizenship', 'Российская Федерация'),
            'CITIZENSHIP': author_data.get('citizenship', 'Российская Федерация'),

            '{{AUTHOR_CITIZENSHIP}}': author_data.get('citizenship', 'Российская Федерация'),
            '{AUTHOR_CITIZENSHIP}': author_data.get('citizenship', 'Российская Федерация'),
            'AUTHOR_CITIZENSHIP': author_data.get('citizenship', 'Российская Федерация'),

            '{{AUTHORS_CITIZENSHIP}}': author_data.get('citizenship', 'Российская Федерация'),
            '{AUTHORS_CITIZENSHIP}': author_data.get('citizenship', 'Российская Федерация'),
            'AUTHORS_CITIZENSHIP': author_data.get('citizenship', 'Российская Федерация'),

            # Адрес (один)
            '{{ADDRESS_FULL}}': author_data.get('address', 'Адрес не указан'),
            '{ADDRESS_FULL}': author_data.get('address', 'Адрес не указан'),
            'ADDRESS_FULL': author_data.get('address', 'Адрес не указан'),

            '{{ADDRESS}}': author_data.get('address', 'Адрес не указан'),
            '{ADDRESS}': author_data.get('address', 'Адрес не указан'),
            'ADDRESS': author_data.get('address', 'Адрес не указан'),

            '{{AUTHOR_ADDRESS}}': author_data.get('address', 'Адрес не указан'),
            '{AUTHOR_ADDRESS}': author_data.get('address', 'Адрес не указан'),
            'AUTHOR_ADDRESS': author_data.get('address', 'Адрес не указан'),

            # Описание вклада (одно, БЕЗ оборудования)
            '{{DESCRIPTION}}': author_data.get('description', 'Вклад не описан'),
            '{DESCRIPTION}': author_data.get('description', 'Вклад не описан'),
            'DESCRIPTION': author_data.get('description', 'Вклад не описан'),

            '{{AUTHOR_DESCRIPTION}}': author_data.get('description', 'Вклад не описан'),
            '{AUTHOR_DESCRIPTION}': author_data.get('description', 'Вклад не описан'),
            'AUTHOR_DESCRIPTION': author_data.get('description', 'Вклад не описан'),

            # Проректор - ИМЯ
            '{{PROVOST_NAME}}': provost_name,
            '{PROVOST_NAME}': provost_name,
            'PROVOST_NAME': provost_name,

            '{{PROVOST}}': provost_name,
            '{PROVOST}': provost_name,
            'PROVOST': provost_name,

            '{{PROVOST_FULL_NAME}}': provost_name,
            '{PROVOST_FULL_NAME}': provost_name,
            'PROVOST_FULL_NAME': provost_name,

            '{{PRO_RECTOR}}': provost_name,
            '{PRO_RECTOR}': provost_name,
            'PRO_RECTOR': provost_name,

            '{{PRO_RECTOR_NAME}}': provost_name,
            '{PRO_RECTOR_NAME}': provost_name,
            'PRO_RECTOR_NAME': provost_name,

            # Должность проректора - ДОБАВЛЕНО
            '{{PROVOST_POST}}': provost_post,
            '{PROVOST_POST}': provost_post,
            'PROVOST_POST': provost_post,

            '{{PRO_RECTOR_POST}}': provost_post,
            '{PRO_RECTOR_POST}': provost_post,
            'PRO_RECTOR_POST': provost_post,

            '{{PROVOST_POSITION}}': provost_post,
            '{PROVOST_POSITION}': provost_post,
            'PROVOST_POSITION': provost_post,

            '{{PRO_RECTOR_POSITION}}': provost_post,
            '{PRO_RECTOR_POSITION}': provost_post,
            'PRO_RECTOR_POSITION': provost_post,

            '{{POST}}': provost_post,
            '{POST}': provost_post,
            'POST': provost_post,

            # Порядковый номер автора (для именования файлов)
            '{{AUTHOR_NUMBER}}': str(author_index + 1),
            '{AUTHOR_NUMBER}': str(author_index + 1),
            'AUTHOR_NUMBER': str(author_index + 1),
        }

        print(f"\nDEBUG: Key replacements for author #{author_index}:")
        print(f"  AUTHOR_FULL_NAME: '{replacements['{{AUTHOR_FULL_NAME}}']}'")
        print(f"  BIRTH_DATE: '{replacements['{{BIRTH_DATE}}']}'")
        print(f"  DAY/MONTH/YEAR: '{replacements['{{DAY}}']}/{replacements['{{MONTH}}']}/{replacements['{{YEAR}}']}'")
        print(f"  CITIZENSHIP: '{replacements['{{CITIZENSHIP}}']}'")
        print(f"  ADDRESS: '{replacements['{{ADDRESS}}']}'")
        print(f"  DESCRIPTION: '{replacements['{{DESCRIPTION}}']}'")
        print(f"  PROGRAM_NAME: '{replacements['{{PROGRAM_NAME}}']}'")
        print(f"  PROVOST_NAME: '{replacements['{{PROVOST_NAME}}']}'")
        print(f"  PROVOST_POST: '{replacements['{{PROVOST_POST}}']}'")
        print(f"  POST: '{replacements['{{POST}}']}'")

        print(f"\n{'=' * 50}")
        print(f"DEBUG: END FILL TEMPLATE FOR AUTHOR #{author_index}")
        print('=' * 50)

        self._replace_in_paragraphs_preserve_formatting(doc, replacements)
        self._replace_in_tables_preserve_formatting(doc, replacements)

    def _extract_single_author_data_detailed(self, author, author_index=0):
        """Извлечение данных одного автора из подробного запроса (14 полей)"""
        result = {
            'full_name': 'Автор не указан',
            'birth_date_full': 'Дата рождения не указана',
            'birth_day': '',
            'birth_day_padded': '',
            'birth_month': '',
            'birth_month_num': '',
            'birth_year': '',
            'citizenship': 'Российская Федерация',
            'address': 'Адрес не указан',
            'description': 'Вклад не описан'
        }

        if not author or not isinstance(author, tuple):
            print(f"DEBUG: Author #{author_index} - Invalid author data")
            return result

        print(f"\nDEBUG: Processing author #{author_index} data")
        print(f"DEBUG: Author tuple length: {len(author)}")

        # 1. ФИО - индексы 7, 8, 9
        if len(author) > 9:
            surname = str(author[7]).strip() if author[7] else ''
            name = str(author[8]).strip() if author[8] else ''
            patronymic = str(author[9]).strip() if author[9] else ''
            full_name = f"{surname} {name} {patronymic}".strip()

            if full_name:
                result['full_name'] = full_name
                print(f"DEBUG: Author #{author_index} - ФИО: '{full_name}'")

        # 2. Дата рождения - индекс 10
        if len(author) > 10 and author[10]:
            birth_date = author[10]
            birth_date_str = ''

            if isinstance(birth_date, datetime):
                birth_date_str = birth_date.strftime('%d.%m.%Y')
            elif isinstance(birth_date, date):
                birth_date_str = birth_date.strftime('%d.%m.%Y')
            elif isinstance(birth_date, str):
                birth_date_str = birth_date.strip()
            else:
                try:
                    birth_date_str = str(birth_date)
                except:
                    birth_date_str = ''

            if birth_date_str:
                result['birth_date_full'] = birth_date_str
                print(f"DEBUG: Author #{author_index} - Date of birth: '{birth_date_str}'")

                # Разбиваем дату на компоненты
                day, month_name, month_num, year = self._parse_birth_date(birth_date_str)

                if day:
                    result['birth_day'] = day
                    result['birth_day_padded'] = day.zfill(2)

                if month_name:
                    result['birth_month'] = month_name

                if month_num:
                    result['birth_month_num'] = month_num.zfill(2)

                if year:
                    result['birth_year'] = year

                print(f"DEBUG: Author #{author_index} - Parsed: Day={day}, Month={month_name}, Year={year}")

        # 3. Гражданство - индекс 11
        if len(author) > 11 and author[11] and isinstance(author[11], str):
            citizenship = str(author[11]).strip()
            if citizenship:
                result['citizenship'] = citizenship
                print(f"DEBUG: Author #{author_index} - Citizenship: '{citizenship}'")

        # 4. Адрес - индекс 12 (целиком)
        if len(author) > 12 and author[12] and isinstance(author[12], str):
            address = str(author[12]).strip()
            if address:
                result['address'] = address
                print(f"DEBUG: Author #{author_index} - Address: '{address}'")

        # 5. Описание вклада - индекс 5 (БЕЗ оборудования)
        if len(author) > 5 and author[5] and isinstance(author[5], str):
            description = str(author[5]).strip()
            if description:
                result['description'] = description
                print(f"DEBUG: Author #{author_index} - Description: '{description}'")

        return result

    def _extract_provost_post(self, settings_data):
        """Извлечение должности проректора из настроек"""
        if not settings_data or not isinstance(settings_data, tuple):
            print("DEBUG: No settings data or not tuple for provost post")
            return 'Проректор по науке и инновациям'  # Значение по умолчанию

        print(f"DEBUG: Settings data tuple length for provost post: {len(settings_data)}")

        # Проверяем возможные индексы для должности проректора
        check_indices = [6, 11, 8]  # Из ваших данных: [6] 'Проректор по науке и инновациям '

        for i in check_indices:
            if len(settings_data) > i:
                item = settings_data[i]
                if isinstance(item, str) and item.strip():
                    post = item.strip()
                    print(f"DEBUG: Using provost post from index {i}: '{post}'")
                    return post

        print("DEBUG: No provost post found in settings, using default")
        return 'Проректор по науке и инновациям'  # Значение по умолчанию

    # Остальные методы остаются без изменений
    def _extract_provost_name(self, settings_data):
        """Извлечение имени проректора из настроек"""
        if not settings_data or not isinstance(settings_data, tuple):
            print("DEBUG: No settings data or not tuple")
            return 'И.И. Иванов'  # Значение по умолчанию

        print(f"DEBUG: Settings data tuple length: {len(settings_data)}")
        print(f"DEBUG: Settings data preview: {settings_data[:5]}...")  # Первые 5 элементов

        # Ищем имя проректора на разных позициях
        provost_name = ''

        # Если не нашли, проверяем конкретные индексы для ИМЕНИ
        check_indices = [10, 5, 9]  # Из ваших данных: [10] 'А.В. Башкиров', [5] 'Башкиров Алексей Викторович'

        for i in check_indices:
            if len(settings_data) > i:
                item = settings_data[i]
                if isinstance(item, str) and item.strip():
                    provost_name = item.strip()
                    print(f"DEBUG: Using provost name from index {i}: '{provost_name}'")
                    return provost_name

        print("DEBUG: No provost name found in settings")
        return 'И.И. Иванов'  # Значение по умолчанию

    def _extract_application_data(self, application):
        """Извлечение данных из заявки"""
        if not application or not isinstance(application, tuple):
            return {'product_name': ''}

        result = {'product_name': ''}

        # Название продукта - индекс 7
        if len(application) > 8 and application[8] and isinstance(application[7], str):
            result['product_name'] = str(application[8]).strip()
            print(f"DEBUG: Found product name at index 7: '{result['product_name']}'")
        elif len(application) > 1 and application[1] and isinstance(application[1], str):
            # Или название заявки - индекс 1
            result['product_name'] = str(application[1]).strip()
            print(f"DEBUG: Found application title at index 1: '{result['product_name']}'")

        return result

    def _parse_birth_date(self, birth_date_str):
        """Парсинг даты рождения на компоненты"""
        day = ''
        month_name = ''
        month_num = ''
        year = ''

        if not birth_date_str:
            return day, month_name, month_num, year

        try:
            # Формат DD.MM.YYYY
            if '.' in birth_date_str:
                parts = birth_date_str.split('.')
                if len(parts) >= 3:
                    day = parts[0].lstrip('0')
                    if not day:
                        day = "0"

                    month_num = parts[1]
                    month_int = int(month_num)
                    month_name = self._get_month_name_ru(month_int)

                    year = parts[2]

            # Формат YYYY-MM-DD
            elif '-' in birth_date_str:
                parts = birth_date_str.split('-')
                if len(parts) >= 3:
                    year = parts[0]

                    month_num = parts[1]
                    month_int = int(month_num)
                    month_name = self._get_month_name_ru(month_int)

                    day = parts[2].lstrip('0')
                    if not day:
                        day = "0"

            elif ' ' in birth_date_str and len(birth_date_str) > 8:
                try:
                    for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%Y/%m/%d']:
                        try:
                            dt = datetime.strptime(birth_date_str, fmt)
                            day = str(dt.day).lstrip('0')
                            if not day:
                                day = "0"
                            month_num = str(dt.month)
                            month_name = self._get_month_name_ru(dt.month)
                            year = str(dt.year)
                            break
                        except ValueError:
                            continue
                except Exception as e:
                    print(f"DEBUG: Error parsing date string '{birth_date_str}': {e}")

        except Exception as e:
            print(f"DEBUG: Error in _parse_birth_date for '{birth_date_str}': {e}")

        return day, month_name, month_num, year

    def _get_month_name_ru(self, month_num):
        """Название месяца на русском в именительном падеже"""
        try:
            months = {
                1: 'январь',
                2: 'февраль',
                3: 'март',
                4: 'апрель',
                5: 'май',
                6: 'июнь',
                7: 'июля',
                8: 'август',
                9: 'сентябрь',
                10: 'октябрь',
                11: 'ноябрь',
                12: 'декабрь'
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
