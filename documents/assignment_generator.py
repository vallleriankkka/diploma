import os
import io
import re
from datetime import datetime, date
from decimal import Decimal
from docx import Document
from flask import current_app


class AssignmentAgreementGenerator:
    """Генератор договора об отчуждении исключительного права - ДЛЯ ВСЕХ АВТОРОВ"""

    def __init__(self, template_name='assignment_agreement_template.docx', debug=True):
        self.debug = debug
        self.template_name = template_name
        self.template_path = self._get_template_path()
        self.setting_dao = None
        self.author_dao = None
        self._log_debug(f"AssignmentAgreementGenerator initialized with template: {template_name}")

    def _log_debug(self, message, data=None):
        """Логирование отладочных сообщений"""
        if self.debug:
            print(f"[DEBUG] {message}")
            if data is not None:
                print(f"[DEBUG] Data: {data}")

    def _get_template_path(self):
        """Получение пути к шаблону"""
        template_path = os.path.join(
            current_app.root_path,
            'templates',
            'documents',
            self.template_name
        )
        self._log_debug(f"Template path: {template_path}")
        self._log_debug(f"Template exists: {os.path.exists(template_path)}")
        return template_path

    def set_daos(self, setting_dao, author_dao=None):
        """Установка DAO объектов для доступа к данным"""
        self.setting_dao = setting_dao
        self.author_dao = author_dao
        self._log_debug(f"DAOs set - setting_dao: {setting_dao is not None}, author_dao: {author_dao is not None}")

    def generate_for_all_authors(self, application, authors_data):
        """Генерация договоров для ВСЕХ авторов"""
        self._log_debug("=" * 60)
        self._log_debug("START generate_for_all_authors")
        self._log_debug("=" * 60)

        try:
            if not os.path.exists(self.template_path):
                self._log_debug(f"ERROR: Template not found: {self.template_path}")
                return []

            if not authors_data:
                self._log_debug("ERROR: No authors data")
                return []

            settings_data = self._get_settings_data()
            if not settings_data:
                self._log_debug("ERROR: Failed to get settings data")
                return []

            self._log_debug(f"Settings data length: {len(settings_data) if settings_data else 0}")
            self._log_debug(f"Settings data: {settings_data}")

            base_amount = self._get_base_amount(settings_data)
            self._log_debug(f"Base amount: {base_amount}")

            documents = []
            self._log_debug(f"Processing {len(authors_data)} authors...")

            for author_index, author in enumerate(authors_data):
                self._log_debug("-" * 40)
                self._log_debug(f"Processing author #{author_index}")

                if not author:
                    self._log_debug(f"  Author #{author_index} - empty data, skipping")
                    continue

                self._log_debug(f"  Author tuple length: {len(author)}")
                self._log_debug(f"  Author raw data: {author}")

                doc_buffer = self._generate_single_author(
                    application, author, authors_data, settings_data,
                    author_index, base_amount
                )

                if doc_buffer:
                    documents.append(doc_buffer)
                    self._log_debug(f"  Document generated successfully")
                else:
                    self._log_debug(f"  ERROR: Failed to generate document")

            self._log_debug("=" * 60)
            self._log_debug(f"END generate_for_all_authors - Generated {len(documents)} documents")
            self._log_debug("=" * 60)
            return documents

        except Exception as e:
            self._log_debug(f"ERROR in generate_for_all_authors: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _generate_single_author(self, application, author, all_authors, settings_data, author_index,
                                author_amount=None):
        """Генерация договора для одного автора"""
        self._log_debug(f"  _generate_single_author for author #{author_index}")

        try:
            if not os.path.exists(self.template_path):
                self._log_debug(f"    ERROR: Template not found: {self.template_path}")
                return None

            if not author:
                self._log_debug(f"    ERROR: No author data for #{author_index}")
                return None

            doc = Document(self.template_path)
            self._log_debug(f"    Document loaded successfully")

            # Отладка шаблона
            self._debug_template_placeholders(doc)

            self._fill_template_for_single_author(
                doc, application, author, all_authors, settings_data,
                author_index, author_amount
            )

            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            self._log_debug(f"    Document saved to buffer, size: {doc_buffer.getbuffer().nbytes} bytes")

            return doc_buffer

        except Exception as e:
            self._log_debug(f"    ERROR in _generate_single_author: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _debug_template_placeholders(self, doc):
        """Отладка: вывод всех плейсхолдеров в шаблоне"""
        self._log_debug("    === DEBUG TEMPLATE PLACEHOLDERS ===")

        placeholders = set()

        # Поиск в параграфах
        for paragraph in doc.paragraphs:
            text = paragraph.text
            # Ищем слова в фигурных скобках
            found = re.findall(r'\{[^}]+\}', text)
            for f in found:
                placeholders.add(f)
            # Ищем слова в двойных фигурных скобках
            found = re.findall(r'\{\{[^}]+\}\}', text)
            for f in found:
                placeholders.add(f)

        # Поиск в таблицах
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text = paragraph.text
                        found = re.findall(r'\{[^}]+\}', text)
                        for f in found:
                            placeholders.add(f)
                        found = re.findall(r'\{\{[^}]+\}\}', text)
                        for f in found:
                            placeholders.add(f)

        self._log_debug(f"    Found placeholders: {sorted(placeholders)}")
        self._log_debug("    === END DEBUG TEMPLATE PLACEHOLDERS ===")

    def _fill_template_for_single_author(self, doc, application, author, all_authors, settings_data, author_index,
                                         author_amount=None):
        """Заполнение шаблона данными для одного автора"""
        self._log_debug(f"    _fill_template_for_single_author for author #{author_index}")
        self._log_debug(f"    Author amount: {author_amount}")

        app_data = self._extract_application_data(application)
        self._log_debug(f"    Application data: {app_data}")

        author_data = self._extract_author_data(author, author_index)
        self._log_debug(f"    Author data extracted:")
        self._log_debug(f"      Full name: {author_data.get('full_name')}")
        self._log_debug(f"      Position: {author_data.get('position')}")
        self._log_debug(f"      Department: {author_data.get('department')}")
        self._log_debug(f"      SNILS: {author_data.get('snils')}")
        self._log_debug(f"      INN: {author_data.get('inn')}")
        self._log_debug(f"      Address: {author_data.get('address')}")
        self._log_debug(f"      Birth day: {author_data.get('birth_day')}")
        self._log_debug(f"      Birth month: {author_data.get('birth_month')}")
        self._log_debug(f"      Birth year: {author_data.get('birth_year')}")
        self._log_debug(f"      Passport series: {author_data.get('passport_series')}")
        self._log_debug(f"      Passport number: {author_data.get('passport_number')}")
        self._log_debug(f"      Passport date: {author_data.get('passport_date')}")
        self._log_debug(f"      Passport by: {author_data.get('passport_by')}")
        self._log_debug(f"      Passport code: {author_data.get('passport_code')}")

        self._log_debug(f"    Extracting provost data from settings...")
        provost1_data = self._extract_provost1_data(settings_data)
        self._log_debug(f"      Provost1 data: {provost1_data}")

        power_of_attorney_data = self._extract_power_of_attorney_data(settings_data)
        self._log_debug(f"      Power of attorney data: {power_of_attorney_data}")

        sum_data = {'value': author_amount} if author_amount is not None else self._extract_sum_data(settings_data)
        self._log_debug(f"    Sum data: {sum_data}")

        other_authors = self._get_other_authors_formatted(all_authors, author, author_index)
        self._log_debug(f"    Other authors: '{other_authors}'")

        power_of_attorney = self._format_power_of_attorney(power_of_attorney_data)
        self._log_debug(f"    Formatted POWER_OF_ATTORNEY: '{power_of_attorney}'")

        sum_amount = self._format_sum(sum_data)
        self._log_debug(f"    Formatted SUM: '{sum_amount}'")

        # Подготовка данных для замены
        provost1_post = provost1_data.get('post', '')
        provost1_post_gen = provost1_data.get('post_genitive', '')
        provost1_name = provost1_data.get('name', '')
        provost1_name_gen = provost1_data.get('name_genitive', '')

        self._log_debug(f"    Provost1 post: '{provost1_post}'")
        self._log_debug(f"    Provost1 post genitive: '{provost1_post_gen}'")
        self._log_debug(f"    Provost1 name: '{provost1_name}'")
        self._log_debug(f"    Provost1 name genitive: '{provost1_name_gen}'")

        # Словарь замен со всеми возможными вариантами
        replacements = {
            # ФИО автора
            'FULLNAME': author_data.get('full_name', 'Автор не указан'),
            'POSITION': author_data.get('position', ''),
            'POSGEN': author_data.get('position_genitive', ''),
            'DEPARTMENT': author_data.get('department', ''),
            'DEPGEN': author_data.get('department_genitive', ''),
            'SNILS': author_data.get('snils', ''),
            'INN': author_data.get('inn', ''),
            'DAY': author_data.get('birth_day', ''),
            'MONTH': author_data.get('birth_month', ''),
            'YEAR': author_data.get('birth_year', ''),
            'SERIES': author_data.get('passport_series', ''),
            'NUMBER': author_data.get('passport_number', ''),
            'DATEP': author_data.get('passport_date', ''),
            'BY': author_data.get('passport_by', ''),
            'CODE': author_data.get('passport_code', ''),
            'ADDRESS': author_data.get('address', 'Адрес не указан'),
            'PRODUCTNAME': app_data.get('product_name', 'Не указано'),
            'POWER': power_of_attorney,
            'SOA': other_authors,
            'SUM': sum_amount,

            # Проректор - разные варианты
            'POSTA': provost1_post_gen,
            'PROVOST': provost1_name_gen,

        }

        # Добавляем варианты с фигурными скобками
        extra_replacements = {}
        for key, value in replacements.items():
            extra_replacements[f'{{{key}}}'] = value
            extra_replacements[f'{{{{ {key} }}}}'] = value
            extra_replacements[f'{{{{{key}}}}}'] = value
            extra_replacements[f'{{{{ {key} }}}}'] = value

        replacements.update(extra_replacements)

        self._log_debug(f"    Key replacements:")
        for key in ['FULLNAME', 'SUM', 'PRODUCTNAME', 'PROVOST1', 'POSTA', 'PROVOST_NAME', 'PROVOST_POST']:
            self._log_debug(f"      {key}: '{replacements.get(key, 'NOT FOUND')}'")

        self._log_debug(f"    Applying replacements in paragraphs...")
        self._replace_in_paragraphs_preserve_formatting(doc, replacements)

        self._log_debug(f"    Applying replacements in tables...")
        self._replace_in_tables_preserve_formatting(doc, replacements)

        self._log_debug(f"    Template filling completed")

    def _get_base_amount(self, settings_data):
        """Получение базовой суммы из настроек"""
        self._log_debug(f"    _get_base_amount called")

        try:
            if settings_data and len(settings_data) > 6 and settings_data[6]:
                try:
                    amount = float(settings_data[6])
                    self._log_debug(f"      Amount from settings[6]: {amount}")
                    return amount
                except Exception as e:
                    self._log_debug(f"      Error parsing settings amount: {e}")

            self._log_debug(f"      Using default amount: 0.00")
            return 0.00

        except Exception as e:
            self._log_debug(f"      ERROR in _get_base_amount: {e}")
            return 0.00

    def _format_sum(self, sum_data):
        """Форматирование суммы договора"""
        self._log_debug(f"    _format_sum: sum_data={sum_data}")

        if not sum_data or not sum_data.get('value'):
            self._log_debug(f"      No sum data, returning placeholder")
            return "___________ (___________________) рублей"

        try:
            value_str = str(sum_data['value']).strip()
            self._log_debug(f"      Value string: '{value_str}'")

            if '.' in value_str or ',' in value_str:
                value_clean = value_str.replace(',', '.')
                num_value = float(value_clean)
                rubles = int(round(num_value))
            else:
                rubles = int(value_str)

            self._log_debug(f"      Rubles: {rubles}")

            num_part = f"{rubles:,}".replace(',', ' ')
            sum_in_words = self._number_to_words(rubles)
            self._log_debug(f"      Number part: '{num_part}', Words: '{sum_in_words}'")

            if 10 <= rubles % 100 <= 20:
                rub_form = 'рублей'
            elif rubles % 10 == 1:
                rub_form = 'рубль'
            elif 2 <= rubles % 10 <= 4:
                rub_form = 'рубля'
            else:
                rub_form = 'рублей'

            result = f"{num_part} ({sum_in_words} {rub_form}) рублей"
            self._log_debug(f"      Result: '{result}'")
            return result

        except Exception as e:
            self._log_debug(f"      ERROR in _format_sum: {e}")
            return "___________ (___________________) рублей"

    def _get_settings_data(self):
        """Получение данных настроек из базы"""
        self._log_debug(f"    _get_settings_data called")

        try:
            if self.setting_dao:
                self._log_debug(f"      Getting settings from DAO...")
                settings = self.setting_dao.get_current()

                if settings:
                    self._log_debug(f"      Settings retrieved successfully, length: {len(settings)}")
                    self._log_debug(f"      Settings type: {type(settings)}")
                    return settings
                else:
                    self._log_debug(f"      No settings found in database")
                    return None
            else:
                self._log_debug(f"      setting_dao is None")
                return None

        except Exception as e:
            self._log_debug(f"      ERROR in _get_settings_data: {e}")
            return None

    def _extract_provost1_data(self, settings_data):
        """Извлечение данных проректора из настроек"""
        self._log_debug(f"      _extract_provost1_data called")

        result = {
            'post': '',
            'post_genitive': '',
            'name': '',
            'name_genitive': '',
        }

        if not settings_data:
            self._log_debug(f"        No settings data")
            return result

        try:
            # Выводим все индексы для отладки
            self._log_debug(f"        Settings data length: {len(settings_data)}")
            for i, val in enumerate(settings_data):
                self._log_debug(f"        settings_data[{i}] = {val} (type: {type(val)})")

            # Индекс 7 - имя проректора
            if len(settings_data) > 7 and settings_data[7]:
                result['name'] = str(settings_data[7]).strip()
                result['name_genitive'] = self._to_genitive_name_full(result['name'])
                self._log_debug(f"        Found provost1 name at index 7: '{result['name']}'")
                self._log_debug(f"        Genitive form: '{result['name_genitive']}'")
            else:
                self._log_debug(f"        No provost1 name found at index 7")
                self._log_debug(f"        settings_data length is {len(settings_data)}, need index 7")

            # Индекс 8 - должность проректора
            if len(settings_data) > 8 and settings_data[8]:
                result['post'] = str(settings_data[8]).strip()
                result['post_genitive'] = self._to_genitive_position(result['post'])
                self._log_debug(f"        Found provost1 post at index 8: '{result['post']}'")
                self._log_debug(f"        Genitive form: '{result['post_genitive']}'")
            else:
                self._log_debug(f"        No provost1 post found at index 8")
                self._log_debug(f"        settings_data length is {len(settings_data)}, need index 8")

        except Exception as e:
            self._log_debug(f"        ERROR in _extract_provost1_data: {e}")
            import traceback
            traceback.print_exc()

        self._log_debug(f"        Final result: {result}")
        return result

    def _extract_power_of_attorney_data(self, settings_data):
        """Извлечение данных доверенности из настроек"""
        self._log_debug(f"      _extract_power_of_attorney_data called")

        result = {
            'number': '',
            'date': '',
            'date_formatted': '',
        }

        if not settings_data:
            self._log_debug(f"        No settings data")
            return result

        try:
            # Индекс 4 - номер доверенности
            if len(settings_data) > 4 and settings_data[4]:
                result['number'] = str(settings_data[4]).strip()
                self._log_debug(f"        Found power of attorney number at index 4: '{result['number']}'")
            else:
                self._log_debug(f"        No power of attorney number at index 4")

            # Индекс 5 - дата доверенности
            if len(settings_data) > 5 and settings_data[5]:
                date_value = settings_data[5]
                self._log_debug(f"        Date value at index 5: {date_value} (type: {type(date_value)})")
                if isinstance(date_value, (datetime, date)):
                    result['date'] = date_value
                    result['date_formatted'] = date_value.strftime('%d.%m.%Y')
                    self._log_debug(
                        f"        Found power of attorney date (datetime) at index 5: {result['date_formatted']}")
                elif isinstance(date_value, str):
                    result['date'] = date_value
                    result['date_formatted'] = self._parse_date(date_value)
                    self._log_debug(
                        f"        Found power of attorney date (string) at index 5: {result['date_formatted']}")
            else:
                self._log_debug(f"        No power of attorney date at index 5")

        except Exception as e:
            self._log_debug(f"        ERROR in _extract_power_of_attorney_data: {e}")

        return result

    def _extract_sum_data(self, settings_data):
        """Извлечение данных суммы из настроек"""
        self._log_debug(f"      _extract_sum_data called")

        result = {
            'value': '',
            'formatted': '',
        }

        if not settings_data:
            self._log_debug(f"        No settings data")
            return result

        try:
            # Индекс 6 - сумма
            if len(settings_data) > 6 and settings_data[6]:
                sum_value = settings_data[6]
                result['value'] = str(sum_value)
                result['formatted'] = self._format_currency(sum_value)
                self._log_debug(f"        Found sum at index 6: '{result['value']}'")
            else:
                self._log_debug(f"        No sum at index 6")

        except Exception as e:
            self._log_debug(f"        ERROR in _extract_sum_data: {e}")

        return result

    def _extract_author_data(self, author, author_index=0):
        """Извлечение данных одного автора из кортежа"""
        self._log_debug(f"      _extract_author_data for author #{author_index}")

        result = {
            'full_name': 'Автор не указан',
            'position': '',
            'department': '',
            'position_genitive': '',
            'department_genitive': '',
            'snils': '',
            'inn': '',
            'birth_day': '',
            'birth_month': '',
            'birth_year': '',
            'passport_series': '',
            'passport_number': '',
            'passport_date': '',
            'passport_by': '',
            'passport_code': '',
            'address': 'Адрес не указан',
        }

        if not author or not isinstance(author, tuple):
            self._log_debug(f"        Invalid author data or not a tuple")
            return result

        self._log_debug(f"        Author tuple length: {len(author)}")
        for i, val in enumerate(author):
            self._log_debug(f"        author[{i}] = {val} (type: {type(val)})")

        try:
            # ФИО из индексов 7, 8, 9
            if len(author) > 9:
                surname = str(author[7]).strip() if author[7] else ''
                name = str(author[8]).strip() if author[8] else ''
                patronymic = str(author[9]).strip() if author[9] else ''
                result['full_name'] = f"{surname} {name} {patronymic}".strip()
                self._log_debug(f"        Full name: '{result['full_name']}'")
            else:
                self._log_debug(f"        Author tuple too short for full name (need >9, have {len(author)})")

            # Должность из индекса 14
            if len(author) > 14 and author[14]:
                result['position'] = str(author[14]).strip()
                result['position_genitive'] = self._to_genitive_position(result['position'])
                self._log_debug(f"        Position: '{result['position']}'")
                self._log_debug(f"        Position genitive: '{result['position_genitive']}'")
            else:
                self._log_debug(f"        No position at index 14")

            # Подразделение из индекса 18
            if len(author) > 18 and author[18]:
                result['department'] = str(author[18]).strip()
                result['department_genitive'] = self._to_genitive_department(result['department'])
                self._log_debug(f"        Department: '{result['department']}'")
                self._log_debug(f"        Department genitive: '{result['department_genitive']}'")
            else:
                self._log_debug(f"        No department at index 18")

            # Адрес из индекса 12
            if len(author) > 12 and author[12]:
                result['address'] = str(author[12]).strip()
                self._log_debug(f"        Address: '{result['address'][:50]}'")
            else:
                self._log_debug(f"        No address at index 12")

            # Дата рождения из индекса 10
            if len(author) > 10 and author[10]:
                birth_date = author[10]
                self._log_debug(f"        Birth date raw: {birth_date} (type: {type(birth_date)})")
                if isinstance(birth_date, (datetime, date)):
                    birth_date_str = birth_date.strftime('%d.%m.%Y')
                else:
                    birth_date_str = str(birth_date)

                day, month_name, month_num, year = self._parse_birth_date(birth_date_str)
                result['birth_day'] = day
                result['birth_month'] = month_name
                result['birth_year'] = year
                self._log_debug(f"        Birth date parsed: {day}.{month_name}.{year}")
            else:
                self._log_debug(f"        No birth date at index 10")

            # Получение паспортных данных из DAO
            if len(author) > 2 and author[2] and self.author_dao:
                try:
                    author_id = int(author[2])
                    self._log_debug(f"        Getting author details from DAO for id={author_id}")
                    author_details = self.author_dao.get_by_id(author_id)

                    if author_details:
                        self._log_debug(f"        Author details retrieved, length: {len(author_details)}")
                        for i, val in enumerate(author_details):
                            self._log_debug(f"        author_details[{i}] = {val}")

                        # Паспорт (индекс 6)
                        if len(author_details) > 6 and author_details[6]:
                            passport_str = str(author_details[6]).strip()
                            passport_clean = ''.join(filter(str.isdigit, passport_str))
                            if len(passport_clean) >= 10:
                                result['passport_series'] = f"{passport_clean[:2]} {passport_clean[2:4]}"
                                result['passport_number'] = passport_clean[4:10]
                                self._log_debug(
                                    f"        Passport series: '{result['passport_series']}', number: '{result['passport_number']}'")

                        # СНИЛС (индекс 7)
                        if len(author_details) > 7 and author_details[7]:
                            snils = str(author_details[7]).strip()
                            digits = ''.join(filter(str.isdigit, snils))
                            if len(digits) == 11:
                                result['snils'] = f"{digits[:3]}-{digits[3:6]}-{digits[6:9]} {digits[9:]}"
                            else:
                                result['snils'] = snils
                            self._log_debug(f"        SNILS: '{result['snils']}'")

                        # ИНН (индекс 8)
                        if len(author_details) > 8 and author_details[8]:
                            result['inn'] = str(author_details[8]).strip()
                            self._log_debug(f"        INN: '{result['inn']}'")

                        # Дата выдачи паспорта (индекс 15)
                        if len(author_details) > 15 and author_details[15]:
                            issue_date = author_details[15]
                            if isinstance(issue_date, (datetime, date)):
                                result['passport_date'] = issue_date.strftime('%d.%m.%Y')
                            else:
                                result['passport_date'] = str(issue_date)
                            self._log_debug(f"        Passport issue date: '{result['passport_date']}'")

                        # Кем выдан (индекс 16)
                        if len(author_details) > 16 and author_details[16]:
                            result['passport_by'] = str(author_details[16]).strip()
                            self._log_debug(f"        Passport issued by: '{result['passport_by'][:50]}'")

                        # Код подразделения (индекс 17)
                        if len(author_details) > 17 and author_details[17]:
                            code = str(author_details[17]).strip()
                            digits = ''.join(filter(str.isdigit, code))
                            if len(digits) == 6:
                                result['passport_code'] = f"{digits[:3]}-{digits[3:]}"
                            else:
                                result['passport_code'] = code
                            self._log_debug(f"        Passport code: '{result['passport_code']}'")

                except Exception as e:
                    self._log_debug(f"        Error getting author details: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                self._log_debug(f"        No author_id or author_dao available")

        except Exception as e:
            self._log_debug(f"        ERROR in _extract_author_data: {e}")
            import traceback
            traceback.print_exc()

        # Заполняем пустые поля
        if not result['passport_series']:
            result['passport_series'] = '__ __'
        if not result['passport_number']:
            result['passport_number'] = '______'
        if not result['passport_date']:
            result['passport_date'] = '__.__.____'
        if not result['passport_by']:
            result['passport_by'] = '________________________________________________________'
        if not result['passport_code']:
            result['passport_code'] = '___-___'
        if not result['snils']:
            result['snils'] = '___-___-___ __'
        if not result['inn']:
            result['inn'] = '__________'

        return result

    def _extract_application_data(self, application):
        """Извлечение данных заявки"""
        result = {'product_name': 'Не указано'}

        if not application:
            self._log_debug(f"        No application data")
            return result

        try:
            if isinstance(application, dict):
                result['product_name'] = application.get('product_name', 'Не указано')
                self._log_debug(f"        Product name from dict: '{result['product_name']}'")
            elif isinstance(application, tuple):
                self._log_debug(f"        Application tuple length: {len(application)}")
                if len(application) > 8 and application[8]:
                    result['product_name'] = str(application[8]).strip()
                    self._log_debug(f"        Product name from tuple[8]: '{result['product_name']}'")
                elif len(application) > 3 and application[3]:
                    result['product_name'] = str(application[3]).strip()
                    self._log_debug(f"        Product name from tuple[3]: '{result['product_name']}'")
        except Exception as e:
            self._log_debug(f"        Error extracting application data: {e}")

        return result

    def _format_power_of_attorney(self, power_of_attorney_data):
        """Форматирование доверенности"""
        if not power_of_attorney_data:
            return ""

        number = power_of_attorney_data.get('number', '').strip()
        date_formatted = power_of_attorney_data.get('date_formatted', '').strip()

        result_parts = []
        if number:
            if not number.startswith('№'):
                result_parts.append(f"№ {number}")
            else:
                result_parts.append(number)
        if date_formatted:
            result_parts.append(f"от {date_formatted}")

        result = " ".join(result_parts) if result_parts else ""
        self._log_debug(f"        Formatted power of attorney: '{result}'")
        return result

    def _number_to_words(self, num):
        """Преобразование числа в пропись"""
        units_female = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
        units_male = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
        teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать',
                 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать']
        tens = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят',
                'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто']
        hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот',
                    'шестьсот', 'семьсот', 'восемьсот', 'девятьсот']

        def convert_three_digits(n, gender='male'):
            result = []
            h = n // 100
            if h > 0:
                result.append(hundreds[h])
            remainder = n % 100
            if remainder >= 20:
                t = remainder // 10
                u = remainder % 10
                result.append(tens[t])
                if u > 0:
                    result.append(units_female[u] if gender == 'female' else units_male[u])
            elif remainder >= 10:
                result.append(teens[remainder - 10])
            elif remainder > 0:
                result.append(units_female[remainder] if gender == 'female' else units_male[remainder])
            return result

        if num == 0:
            return 'ноль'

        result = []
        billions = num // 1_000_000_000
        if billions > 0:
            result.extend(convert_three_digits(billions))
            if 10 <= billions % 100 <= 20:
                result.append('миллиардов')
            elif billions % 10 == 1:
                result.append('миллиард')
            elif 2 <= billions % 10 <= 4:
                result.append('миллиарда')
            else:
                result.append('миллиардов')

        millions = (num % 1_000_000_000) // 1_000_000
        if millions > 0:
            result.extend(convert_three_digits(millions))
            if 10 <= millions % 100 <= 20:
                result.append('миллионов')
            elif millions % 10 == 1:
                result.append('миллион')
            elif 2 <= millions % 10 <= 4:
                result.append('миллиона')
            else:
                result.append('миллионов')

        thousands = (num % 1_000_000) // 1000
        if thousands > 0:
            result.extend(convert_three_digits(thousands, gender='female'))
            if 10 <= thousands % 100 <= 20:
                result.append('тысяч')
            elif thousands % 10 == 1:
                result.append('тысяча')
            elif 2 <= thousands % 10 <= 4:
                result.append('тысячи')
            else:
                result.append('тысяч')

        units = num % 1000
        if units > 0 or num == 0:
            result.extend(convert_three_digits(units))

        return ' '.join([w for w in result if w])

    def _to_genitive_name_full(self, name):
        """Преобразование полного имени в родительный падеж"""
        if not name:
            return ''
        name = name.strip()
        parts = name.split()
        if len(parts) < 2:
            return self._to_genitive_name(name)

        genitive_parts = []
        for i, part in enumerate(parts):
            part_lower = part.lower()
            if i == 0:  # Фамилия
                if part_lower.endswith(('ов', 'ев', 'ин')):
                    gen_part = part + 'а'
                elif part_lower.endswith(('ский', 'цкий', 'ой', 'ий', 'ый')):
                    gen_part = part[:-2] + 'ого'
                else:
                    gen_part = part + 'а'
            else:  # Имя и отчество
                # Специальные правила для имен
                if part_lower == 'алексей':
                    gen_part = 'Алексея'
                elif part_lower == 'дмитрий':
                    gen_part = 'Дмитрия'
                elif part_lower == 'сергей':
                    gen_part = 'Сергея'
                elif part_lower == 'андрей':
                    gen_part = 'Андрея'
                elif part_lower == 'николай':
                    gen_part = 'Николая'
                elif part_lower.endswith('ий'):
                    gen_part = part[:-2] + 'ия'
                elif part_lower.endswith('ей'):
                    gen_part = part[:-2] + 'я'
                elif part_lower.endswith('я'):
                    gen_part = part[:-1] + 'и'
                elif part_lower.endswith('а'):
                    gen_part = part[:-1] + 'ы'
                elif part_lower.endswith('ь'):
                    gen_part = part[:-1] + 'я'
                elif part_lower.endswith('й'):
                    gen_part = part[:-1] + 'я'
                else:
                    gen_part = part + 'а'

            if part and part[0].isupper() and gen_part:
                gen_part = gen_part[0].upper() + gen_part[1:]
            genitive_parts.append(gen_part)

        return ' '.join(genitive_parts)

    def _to_genitive_name(self, name):
        """Преобразование фамилии в родительный падеж"""
        if not name:
            return ''
        name = name.strip()
        parts = name.split()
        if not parts:
            return name

        last_name = parts[0]
        last_name_lower = last_name.lower()

        if last_name_lower.endswith(('ов', 'ев', 'ин')):
            last_name_gen = last_name + 'а'
        elif last_name_lower.endswith(('ский', 'цкий', 'ой', 'ий', 'ый')):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('ая'):
            last_name_gen = last_name[:-2] + 'ой'
        elif last_name_lower.endswith('яя'):
            last_name_gen = last_name[:-2] + 'ей'
        else:
            last_name_gen = last_name + 'а'

        if last_name and last_name[0].isupper() and last_name_gen:
            last_name_gen = last_name_gen[0].upper() + last_name_gen[1:]

        return last_name_gen if len(parts) == 1 else ' '.join([last_name_gen] + parts[1:])

    def _to_genitive_position(self, position):
        """Преобразование должности в родительный падеж"""
        if not position:
            return ''

        position_lower = position.lower().strip()

        special_cases = {
            'проректор по науке и инновациям': 'Проректора по науке и инновации',
            'ректор': 'Ректора',
            'директор': 'Директора',
            'заведующий кафедрой': 'Заведующего кафедрой',
            'начальник отдела': 'Начальника отдела',
            'руководитель': 'Руководителя'
        }

        if position_lower in special_cases:
            return special_cases[position_lower]

        words = position.split()
        if not words:
            return position

        first_word = words[0]
        first_word_lower = first_word.lower()

        if first_word_lower.endswith(('ор', 'ист', 'ер')):
            first_word_gen = first_word + 'а'
        elif first_word_lower.endswith('ль'):
            if first_word_lower.endswith('ель'):
                first_word_gen = first_word[:-2] + 'ля'
            else:
                first_word_gen = first_word[:-2] + 'ля'
        elif first_word_lower.endswith(('ой', 'ий', 'ый')):
            first_word_gen = first_word[:-2] + 'ого'
        elif first_word_lower.endswith('а'):
            first_word_gen = first_word[:-1] + 'ы'
        elif first_word_lower.endswith('я'):
            first_word_gen = first_word[:-1] + 'и'
        else:
            first_word_gen = first_word + 'а'

        if first_word and first_word[0].isupper() and first_word_gen:
            first_word_gen = first_word_gen[0].upper() + first_word_gen[1:]

        return ' '.join([first_word_gen] + words[1:]) if len(words) > 1 else first_word_gen

    def _to_genitive_department(self, department):
        """Преобразование подразделения в родительный падеж"""
        if not department:
            return ''

        dept = department.strip()
        dept_lower = dept.lower()

        if 'кафедр' in dept_lower:
            if dept_lower.endswith('кафедра'):
                return dept[:-6] + 'кафедры'
            return dept + 'ы'
        elif 'отдел' in dept_lower:
            if dept_lower.endswith('отдел'):
                return dept[:-5] + 'отдела'
            return dept + 'а'
        elif 'лаборатор' in dept_lower:
            if dept_lower.endswith('лаборатория'):
                return dept[:-11] + 'лаборатории'
            return dept + 'ии'
        elif 'факультет' in dept_lower:
            if dept_lower.endswith('факультет'):
                return dept[:-9] + 'факультета'
            return dept + 'а'
        elif 'институт' in dept_lower:
            if dept_lower.endswith('институт'):
                return dept[:-8] + 'института'
            return dept + 'а'
        elif 'центр' in dept_lower:
            if dept_lower.endswith('центр'):
                return dept[:-5] + 'центра'
            return dept + 'а'

        if dept_lower.endswith('ая'):
            return dept[:-2] + 'ой'
        elif dept_lower.endswith('яя'):
            return dept[:-2] + 'ей'
        elif dept_lower.endswith(('ий', 'ый')):
            return dept[:-2] + 'ого'
        elif dept_lower.endswith('а'):
            return dept[:-1] + 'ы'
        elif dept_lower.endswith('я'):
            return dept[:-1] + 'и'
        elif dept_lower.endswith('о'):
            return dept[:-1] + 'а'
        elif dept_lower.endswith('е'):
            return dept[:-1] + 'я'
        elif dept_lower.endswith('ь'):
            return dept[:-1] + 'я'
        else:
            return dept + 'а'

    def _get_month_name_ru(self, month_num):
        """Название месяца на русском"""
        months = {
            1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля',
            5: 'мая', 6: 'июня', 7: 'июля', 8: 'августа',
            9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
        }
        return months.get(int(month_num), '') if month_num else ''

    def _parse_date(self, date_str):
        """Парсинг даты из строки"""
        if not date_str:
            return ''
        try:
            for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y.%m.%d'):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt.strftime('%d.%m.%Y')
                except:
                    continue
            return date_str
        except:
            return date_str

    def _format_currency(self, value):
        """Форматирование денежной суммы"""
        try:
            if isinstance(value, (int, float)):
                num_value = float(value)
            elif isinstance(value, str):
                cleaned = ''.join(c for c in value if c.isdigit() or c in '.,')
                cleaned = cleaned.replace(',', '.')
                num_value = float(cleaned) if cleaned else 0
            else:
                return str(value)

            if num_value.is_integer():
                return f"{int(num_value):,}".replace(',', ' ') + ' рублей'
            else:
                return f"{num_value:,.2f}".replace(',', ' ') + ' рублей'
        except:
            return str(value)

    def _get_other_authors_formatted(self, all_authors, current_author, current_index):
        """Получение списка остальных авторов"""
        if not all_authors or len(all_authors) <= 1:
            return ""

        other_authors = []
        for i, author in enumerate(all_authors):
            if author == current_author or i == current_index:
                continue
            if len(author) > 9:
                surname = str(author[7]).strip() if author[7] else ''
                name = str(author[8]).strip() if author[8] else ''
                patronymic = str(author[9]).strip() if author[9] else ''
                full_name = f"{surname} {name} {patronymic}".strip()
                if full_name and full_name != '  ':
                    other_authors.append(full_name)

        return f"(в соавторстве {', '.join(other_authors)})" if other_authors else ""

    def _parse_birth_date(self, birth_date_str):
        """Парсинг даты рождения"""
        day = ''
        month_name = ''
        month_num = ''
        year = ''

        if not birth_date_str:
            return day, month_name, month_num, year

        try:
            if '.' in birth_date_str:
                parts = birth_date_str.split('.')
                if len(parts) >= 3:
                    day = parts[0].lstrip('0')
                    month_num = parts[1]
                    if month_num:
                        month_name = self._get_month_name_ru(int(month_num))
                    year = parts[2]
            elif '-' in birth_date_str:
                parts = birth_date_str.split('-')
                if len(parts) >= 3:
                    year = parts[0]
                    month_num = parts[1]
                    if month_num:
                        month_name = self._get_month_name_ru(int(month_num))
                    day = parts[2].lstrip('0')
        except:
            pass

        return day, month_name, month_num, year

    def _replace_in_paragraphs_preserve_formatting(self, doc, replacements):
        """Замена в параграфах с объединением runs для корректной замены"""
        replaced_count = 0

        for paragraph in doc.paragraphs:
            # Получаем полный текст параграфа
            full_text = paragraph.text

            # Проверяем, есть ли хоть один плейсхолдер в параграфе
            needs_replace = any(old_text in full_text for old_text in replacements.keys())

            if not needs_replace:
                continue

            # Объединяем все runs в один для этого параграфа
            # (сохраняем стиль первого run)
            if len(paragraph.runs) > 0:
                # Сохраняем стиль первого run
                first_run = paragraph.runs[0]
                font_style = first_run.font
                alignment = paragraph.alignment

                # Получаем весь текст
                full_text = ''.join(run.text for run in paragraph.runs)

                # Применяем все замены
                for old_text, new_text in replacements.items():
                    if old_text in full_text:
                        full_text = full_text.replace(old_text, new_text)
                        replaced_count += 1

                # Очищаем все runs
                for run in paragraph.runs:
                    run.text = ''

                # Устанавливаем новый текст в первый run
                first_run.text = full_text

                # Восстанавливаем стиль
                first_run.font.name = font_style.name
                first_run.font.size = font_style.size
                first_run.font.bold = font_style.bold
                first_run.font.italic = font_style.italic
                paragraph.alignment = alignment

        self._log_debug(f"        Total replacements in paragraphs: {replaced_count}")

    def _replace_in_tables_preserve_formatting(self, doc, replacements):
        """Замена в таблицах с объединением runs для корректной замены"""
        replaced_count = 0

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # Получаем полный текст параграфа
                        full_text = paragraph.text

                        # Проверяем, есть ли хоть один плейсхолдер в параграфе
                        needs_replace = any(old_text in full_text for old_text in replacements.keys())

                        if not needs_replace:
                            continue

                        # Объединяем все runs в один для этого параграфа
                        if len(paragraph.runs) > 0:
                            # Сохраняем стиль первого run
                            first_run = paragraph.runs[0]
                            font_style = first_run.font
                            alignment = paragraph.alignment

                            # Получаем весь текст
                            full_text = ''.join(run.text for run in paragraph.runs)

                            # Применяем все замены
                            for old_text, new_text in replacements.items():
                                if old_text in full_text:
                                    full_text = full_text.replace(old_text, new_text)
                                    replaced_count += 1

                            # Очищаем все runs
                            for run in paragraph.runs:
                                run.text = ''

                            # Устанавливаем новый текст в первый run
                            first_run.text = full_text

                            # Восстанавливаем стиль
                            first_run.font.name = font_style.name
                            first_run.font.size = font_style.size
                            first_run.font.bold = font_style.bold
                            first_run.font.italic = font_style.italic
                            paragraph.alignment = alignment

        self._log_debug(f"        Total replacements in tables: {replaced_count}")
