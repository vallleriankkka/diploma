import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from flask import current_app
import re


class RIDNotificationGenerator:
    """Document generator for RID creation notification"""

    def __init__(self, template_name='notification_template.docx'):
        self.template_name = template_name
        self.template_path = self._get_template_path()
        self.setting_dao = None
        self.application_dao = None

    def set_daos(self, setting_dao, application_dao):
        """Set DAO instances for accessing data"""
        self.setting_dao = setting_dao
        self.application_dao = application_dao

    def _get_template_path(self):
        """Get template path"""
        try:
            if current_app:
                return os.path.join(
                    current_app.root_path,
                    'templates',
                    'documents',
                    self.template_name
                )
        except RuntimeError:
            pass

        return os.path.join('templates', 'documents', self.template_name)

    def _to_dative_case(self, name_parts, position):
        """
        Преобразует ФИО и должность в дательный падеж

        Args:
            name_parts: tuple (фамилия, имя, отчество) или строка с полным ФИО
            position: должность в именительном падеже

        Returns:
            tuple: (должность_в_дательном, ФИО_в_дательном)
        """
        # Правила склонения должностей
        position_dative = position
        position_lower = position.lower()

        # Специальные правила для длинных фраз
        if 'проректор по науке и инновациям' in position_lower:
            position_dative = 'Проректору по науке и инновациям'
        elif 'проректор' in position_lower:
            # Склоняем только слово "проректор"
            if position_lower.endswith('тор'):
                position_dative = position[:-3] + 'тору'
            else:
                position_dative = position + 'у'
            # Добавляем остальную часть должности
            rest = position[len('проректор'):] if position.lower().startswith('проректор') else ''
            if rest:
                position_dative = position_dative + rest
        elif position_lower.endswith('ая'):  # женский род
            position_dative = position[:-2] + 'ой'
        elif position_lower.endswith('ой'):  # мужской род на -ой
            position_dative = position[:-3] + 'ому'
        elif position_lower.endswith('ий'):  # мужской род на -ий
            position_dative = position[:-2] + 'ему'
        elif position_lower.endswith('ый'):  # мужской род на -ый
            position_dative = position[:-2] + 'ому'
        elif position_lower.endswith('рь'):  # на -рь
            position_dative = position + 'ю'
        elif position_lower.endswith('ь'):  # на -ь
            position_dative = position[:-1] + 'ю'
        elif position_lower.endswith('а'):  # женский род на -а
            position_dative = position[:-1] + 'е'
        elif position_lower.endswith('я'):  # женский род на -я
            position_dative = position[:-1] + 'е'
        elif position_lower.endswith('ек'):  # на -ек
            position_dative = position[:-2] + 'ку'
        else:
            position_dative = position + 'у'

        # Обработка ФИО
        if isinstance(name_parts, tuple) and len(name_parts) >= 3:
            surname, name, patronymic = name_parts
        elif isinstance(name_parts, str):
            parts = name_parts.strip().split()
            if len(parts) >= 3:
                surname, name, patronymic = parts[0], parts[1], parts[2]
            elif len(parts) == 2:
                surname, name = parts
                patronymic = ''
            else:
                surname = name_parts
                name = ''
                patronymic = ''
        else:
            return position_dative, name_parts if isinstance(name_parts, str) else ''

        # Склонение фамилии
        surname_dative = surname
        if surname.endswith('ов') or surname.endswith('ев') or surname.endswith('ин'):
            surname_dative = surname + 'у'
        elif surname.endswith('ий'):
            surname_dative = surname[:-2] + 'ему'
        elif surname.endswith('ый'):
            surname_dative = surname[:-2] + 'ому'
        elif surname.endswith('а'):
            surname_dative = surname[:-1] + 'е'
        elif surname.endswith('я'):
            surname_dative = surname[:-1] + 'е'
        elif surname.endswith('ко'):
            surname_dative = surname  # неизменяемые фамилии на -ко
        else:
            surname_dative = surname + 'у'

        # Склонение имени
        name_dative = name
        if name.endswith('й'):
            name_dative = name[:-1] + 'ю'
        elif name.endswith('а'):
            name_dative = name[:-1] + 'е'
        elif name.endswith('я'):
            name_dative = name[:-1] + 'е'
        elif name.endswith('ь'):
            name_dative = name[:-1] + 'ю'
        elif name.endswith('ей'):
            name_dative = name  # Сергей -> Сергею
            if name == 'Сергей':
                name_dative = 'Сергею'

        # Склонение отчества
        patronymic_dative = patronymic
        if patronymic.endswith('на'):
            patronymic_dative = patronymic[:-2] + 'не'
        elif patronymic.endswith('ча'):
            patronymic_dative = patronymic[:-2] + 'чу'
        elif patronymic.endswith('вна'):
            patronymic_dative = patronymic[:-3] + 'вне'
        elif patronymic.endswith('ич'):
            patronymic_dative = patronymic + 'у'
        elif patronymic.endswith('чна'):
            patronymic_dative = patronymic[:-3] + 'чне'

        # Собираем ФИО в дательном падеже
        if name_dative and patronymic_dative:
            full_name_dative = f"{surname_dative} {name_dative} {patronymic_dative}"
        elif name_dative:
            full_name_dative = f"{surname_dative} {name_dative}"
        else:
            full_name_dative = surname_dative

        return position_dative, full_name_dative.strip()

    def generate(self, application_id, user_id):
        """Generate RID notification document"""
        try:
            template_path = self._get_template_path()

            if not os.path.exists(template_path):
                print(f"ERROR: Template not found: {template_path}")
                return None

            print(f"Loading template from: {template_path}")
            doc = Document(template_path)

            # Get all necessary data
            application_data = self._get_application_data(application_id, user_id)
            if not application_data:
                print("ERROR: Application not found")
                return None

            print(f"\n=== APPLICATION DATA STRUCTURE ===")
            print(f"Type: {type(application_data)}")
            print(f"Length: {len(application_data) if hasattr(application_data, '__len__') else 'N/A'}")

            # Выводим все поля с индексами
            if isinstance(application_data, tuple):
                for i, value in enumerate(application_data):
                    print(f"[{i:2d}] '{value}' (type: {type(value).__name__})")
            print("================================\n")

            # Get provost data from settings
            provost_name, provost_post = self._get_provost_data()
            print(f"Provost data: {provost_name}, {provost_post}")

            # Get authors data
            authors_data = self._get_authors_data(application_id)
            print(f"Authors data received: {len(authors_data) if authors_data else 0} authors")

            # Process authors data
            processed_authors = self._process_authors_data_for_doc(authors_data)
            print(f"Processed authors: {len(processed_authors)}")

            self._fill_template(doc, application_data, processed_authors, provost_name, provost_post)
            return doc

        except Exception as e:
            print(f"ERROR in generate: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _get_application_data(self, application_id, user_id):
        """Get application data WITHOUT access check"""
        print(f"\n[RIDNotificationGenerator._get_application_data] NO ACCESS CHECK")
        print(f"  application_id: {application_id}")
        print(f"  user_id: {user_id} (ignored)")

        if not self.application_dao:
            print(f"  ❌ ERROR: application_dao is None!")
            return None

        try:
            # Вариант 1: Используем метод без проверки пользователя
            if hasattr(self.application_dao, 'get_application_by_id'):
                print(f"  Using get_application_by_id({application_id}) - NO USER CHECK")
                application_data = self.application_dao.get_application_by_id(application_id)
            else:
                print(f"  ⚠️ No suitable methods found")
                return None

            if application_data:
                print(f"  ✅ Data received successfully")
                return application_data
            else:
                print(f"  ❌ No data returned (None)")
                return None

        except Exception as e:
            print(f"  ❌ Error getting application data: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _get_authors_data(self, application_id):
        """Get detailed authors data"""
        print(f"\n[_get_authors_data] Получение авторов для заявки {application_id}")

        if not self.application_dao:
            print("  ❌ ERROR: application_dao is None!")
            return []

        try:
            # Используем get_application_authors_detailed
            if hasattr(self.application_dao, 'get_application_authors_detailed'):
                print(f"  Using get_application_authors_detailed({application_id})")
                authors = self.application_dao.get_application_authors_detailed(application_id)
            else:
                print(f"  ⚠️ No author methods found")
                return []

            print(f"  Found {len(authors) if authors else 0} authors")
            return authors or []

        except Exception as e:
            print(f"  ❌ Error getting authors data: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _process_authors_data_for_doc(self, authors_list):
        """Process authors data for document tables - ОСНОВНОЙ МЕТОД"""
        print(f"\n[_process_authors_data_for_doc] Processing {len(authors_list) if authors_list else 0} authors")

        if not authors_list:
            print("  ❌ Authors list is empty!")
            return [self._create_test_author()]

        processed_authors = []

        for idx, author in enumerate(authors_list):
            print(f"\n  --- Processing author {idx + 1} ---")

            try:
                # Обработка tuple автора
                if isinstance(author, tuple):
                    print(f"    Tuple with {len(author)} elements")
                    processed_author = self._process_tuple_author(author, idx)
                else:
                    print(f"    ⚠️ Unknown type: {type(author)}")
                    continue

                if processed_author:
                    processed_authors.append(processed_author)
                    print(f"    ✅ Added: {processed_author['full_name']}")

            except Exception as e:
                print(f"    ❌ Error processing author {idx}: {e}")
                continue

        print(f"\n  ✅ Total processed authors: {len(processed_authors)}")

        # Если авторов нет, добавляем тестового
        if not processed_authors:
            print("  ⚠️ No authors processed, adding test author")
            processed_authors.append(self._create_test_author())

        return processed_authors

    def _process_tuple_author(self, author_tuple, idx):
        """Process author data from tuple"""
        # ИСПРАВЛЕНО: Индексы для данных автора
        # По вашим данным:
        # [7] - фамилия
        # [8] - имя
        # [9] - отчество
        # [3] - процент участия
        # [5] - время затраченное (текстовое поле 'ббе')
        # [6] - оборудование
        # [13] - студент (False/True)
        # [17] - должность
        # [18] - подразделение

        # Базовые данные
        surname = ''
        name = ''
        patronymic = ''
        participation_percent = None
        time_spent = None
        equipment = ''
        is_student = False
        position = ''
        department = ''

        # Получаем данные по индексам
        if len(author_tuple) > 7:
            surname = str(author_tuple[7]) if author_tuple[7] else ''
        if len(author_tuple) > 8:
            name = str(author_tuple[8]) if author_tuple[8] else ''
        if len(author_tuple) > 9:
            patronymic = str(author_tuple[9]) if author_tuple[9] else ''
        if len(author_tuple) > 3:
            participation_percent = author_tuple[3]

        # ВРЕМЯ - ПРАВИЛЬНАЯ ОБРАБОТКА
        if len(author_tuple) > 4:
            time_val = author_tuple[4]
            if time_val:
                # Пробуем извлечь число из строки
                if isinstance(time_val, str):
                    numbers = re.findall(r'\d+', str(time_val))
                    if numbers:
                        time_spent = int(numbers[0])
                    else:
                        # Если нет чисел, используем значение по умолчанию
                        time_spent = 40
                elif isinstance(time_val, (int, float)):
                    time_spent = int(time_val)
                else:
                    time_spent = 40
            else:
                time_spent = 40

        if len(author_tuple) > 6:
            equipment = str(author_tuple[6]) if author_tuple[6] else ''
        if len(author_tuple) > 13:
            is_student = bool(author_tuple[13])
        if len(author_tuple) > 17:
            position = str(author_tuple[17]) if author_tuple[17] else ''
        if len(author_tuple) > 18:
            department = str(author_tuple[18]) if author_tuple[18] else ''

        print(f"      Данные автора:")
        print(f"        Фамилия: '{surname}'")
        print(f"        Имя: '{name}'")
        print(f"        Отчество: '{patronymic}'")
        print(f"        Процент: '{participation_percent}'")
        print(f"        Время: '{time_spent}' (обработано из: '{author_tuple[5] if len(author_tuple) > 5 else ''}')")
        print(f"        Оборудование: '{equipment}'")
        print(f"        Студент: {is_student}")
        print(f"        Должность: '{position}'")
        print(f"        Подразделение: '{department}'")

        return self._create_author_dict(
            surname, name, patronymic,
            participation_percent, time_spent, equipment,
            is_student, '', position, department
        )

    def _create_author_dict(self, surname, name, patronymic,
                            participation_percent, time_spent, equipment,
                            is_student, education_level, position, department):
        """Create standardized author dictionary"""
        # Полное имя
        full_name = f"{surname} {name} {patronymic}".strip()

        # Форматирование процента
        percentage_str = self._format_percentage(participation_percent)

        # Форматирование времени - УЖЕ ПРАВИЛЬНО ОБРАБОТАНО
        time_str = str(time_spent) if time_spent is not None else "40"

        # Статус в родительном падеже
        status_genitive = self._format_author_status(
            is_student, education_level, position, department
        )

        return {
            'full_name': full_name,
            'surname': surname,
            'name': name,
            'patronymic': patronymic,
            'status_genitive': status_genitive if status_genitive else 'сотрудник/студент',
            'percentage': percentage_str,
            'time_spent': time_str,
            'equipment': str(equipment) if equipment else '',
            'is_student': is_student,
            'education_level': education_level,
            'position': position,
            'department': department
        }

    def _create_test_author(self):
        """Create test author for debugging"""
        return {
            'full_name': 'Иванов Иван Иванович',
            'surname': 'Иванов',
            'name': 'Иван',
            'patronymic': 'Иванович',
            'status_genitive': 'аспирант кафедры информатики',
            'percentage': '50%',
            'time_spent': '40',
            'equipment': 'Компьютер №123, лаборатория №1',
            'is_student': True,
            'education_level': 'аспирант',
            'position': '',
            'department': 'кафедра информатики'
        }

    def _format_percentage(self, participation_percent):
        """Format percentage value"""
        if participation_percent is None:
            return "100%"

        try:
            # Преобразуем в строку
            perc_str = str(participation_percent).strip()

            # Убираем лишние символы
            perc_str = perc_str.replace(',', '.').replace('%', '').strip()

            if not perc_str:
                return "100%"

            # Пробуем преобразовать в число
            perc_float = float(perc_str)

            # Если число меньше 1, умножаем на 100 (это дробь)
            if perc_float < 1:
                perc_float = perc_float * 100

            # Форматируем
            if perc_float.is_integer():
                return f"{int(perc_float)}%"
            else:
                formatted = f"{perc_float:.1f}%".replace('.0%', '%')
                return formatted

        except Exception as e:
            print(f"    ⚠️ Error formatting percentage '{participation_percent}': {e}")
            return "100%"

    def _get_provost_data(self):
        """Get provost data in dative case for signatures"""

        if not self.setting_dao:
            print("WARNING: SettingDAO not set, cannot get provost data from database")
            return "", ""

        try:
            settings = self.setting_dao.get_current()
            if settings:
                print(f"DEBUG: Settings received - Type: {type(settings)}, Length: {len(settings)}")

                # Выводим структуру настроек для отладки
                print("\n=== SETTINGS DATA STRUCTURE ===")
                if isinstance(settings, tuple):
                    for i, value in enumerate(settings):
                        print(f"[{i:2d}] '{value}' (type: {type(value).__name__})")
                print("================================\n")

                # Инициализируем переменные
                provost_full_name_nominative = ""
                provost_post_nominative = ""

                # Получаем полное ФИО (индекс 7)
                if len(settings) > 7 and settings[7] and isinstance(settings[7], str):
                    provost_full_name_nominative = settings[7].strip()
                    print(f"  Full name from index 7: '{provost_full_name_nominative}'")
                else:
                    print(f"  WARNING: No full name found at index 7")

                # Получаем должность (индекс 8)
                if len(settings) > 8 and settings[8] and isinstance(settings[8], str):
                    provost_post_nominative = settings[8].strip()
                    print(f"  Position from index 8: '{provost_post_nominative}'")
                else:
                    print(f"  WARNING: No position found at index 8")

                if not provost_full_name_nominative:
                    print(f"  ERROR: No provost full name found in database!")
                    return "", ""

                print(
                    f"Provost (nominative): должность='{provost_post_nominative}', ФИО='{provost_full_name_nominative}'")

                # Преобразуем в дательный падеж
                post_dative, name_dative = self._to_dative_case(provost_full_name_nominative, provost_post_nominative)
                print(f"Provost (dative): должность='{post_dative}', ФИО='{name_dative}'")

                return post_dative, name_dative
            else:
                print(f"  ERROR: No settings found in database!")
                return "", ""

        except Exception as e:
            print(f"Error getting provost data: {e}")
            import traceback
            traceback.print_exc()
            return "", ""

    def _fill_template(self, doc, application, processed_authors, provost_post_dative, provost_name_dative):
        """Fill template with data

        Args:
            provost_post_dative: должность проректора в дательном падеже (например, "Проректору по науке и инновациям")
            provost_name_dative: ФИО проректора в дательном падеже (например, "Башкирову Алексею Викторовичу")
        """
        print("\n" + "=" * 80)
        print("DEBUG: Filling RID notification template")
        print("=" * 80)

        # Extract product name
        product_name = self._extract_product_name(application)
        print(f"Product name: '{product_name}'")

        # Extract NIR info
        nir_info = self._extract_nir_info(application)
        print(f"NIR info: '{nir_info}'")

        # Current date
        current_date = datetime.now().strftime('%d.%m.%Y')
        print(f"Current date: {current_date}")

        # Create replacements - ПРАВИЛЬНО используем аргументы
        replacements = {
            '{{PRODUCT_NAME}}': product_name,
            '{PRODUCT_NAME}': product_name,
            'PRODUCT_NAME': product_name,

            '{{NIR}}': nir_info,
            '{NIR}': nir_info,
            'NIR': nir_info,

            # Должность в дательном (первый аргумент)
            '{{PROVOST_POST_DATIVE}}': provost_post_dative,
            'POST': provost_post_dative,

            # ФИО в дательном (второй аргумент)
            '{{PROVOST_NAME_DATIVE}}': provost_name_dative,
            'PROVOST': provost_name_dative,

            # Полная подпись (должность + ФИО)
            '{{PROVOST_FULL_DATIVE}}': f"{provost_post_dative} {provost_name_dative}",
            '{PROVOST_FULL_DATIVE}': f"{provost_post_dative} {provost_name_dative}",

            '{{CURRENT_DATE}}': current_date,
            '{CURRENT_DATE}': current_date,
            'CURRENT_DATE': current_date,
        }

        print(f"Provost data in dative case:")
        print(f"  Post (dative): '{provost_post_dative}'")
        print(f"  Name (dative): '{provost_name_dative}'")
        print(f"  Full (dative): '{provost_post_dative} {provost_name_dative}'")

        print("\nApplying replacements...")
        self._replace_in_paragraphs(doc, replacements)
        self._replace_in_tables(doc, replacements)

        # Create tables
        self._create_table1(doc, processed_authors)
        self._create_table2(doc, processed_authors)

        # Create signature blocks - на месте SIGNATURE
        self._create_signature_blocks(doc, processed_authors)

        print("\nTemplate filled successfully")
        print("=" * 80)

    def _format_author_status(self, is_student, education_level, position, department):
        """Format author status for table in genitive case"""
        status_parts = []

        if is_student:
            if education_level:
                status_parts.append(education_level)
            else:
                status_parts.append("студент")
        elif position:
            status_parts.append(position)

        if department:
            department_genitive = self._to_genitive_department(department)
            status_parts.append(department_genitive)

        return " ".join(status_parts) if status_parts else "сотрудник/студент"

    def _create_table1(self, doc, authors):
        """Create Table 1 with authors information"""
        print(f"\n[_create_table1] Creating table for {len(authors)} authors")

        # Create table data
        table_data = []
        for author in authors:
            row = [
                author['full_name'],
                author['status_genitive'],
                author['percentage']
            ]
            table_data.append(row)
            print(f"  Row: {row}")

        # Headers
        headers = [
            "ФИО полностью",
            "Должность, место работы\nдля обучающегося – уровень образования (аспирант, магистр, специалист, бакалавр, группа).",
            "Показатель\nтворческого участия, %"
        ]

        # Find and replace placeholder
        success = self._replace_placeholder_with_table(doc, 'TABLE', table_data, headers)

        if not success:
            print("  ⚠️ TABLE placeholder not found, adding table at the end")
            self._add_table_at_end(doc, headers, table_data, "Таблица 1")

    def _create_table2(self, doc, authors):
        """Create Table 2 with equipment and time information"""
        print(f"\n[_create_table2] Creating table for {len(authors)} authors")

        # Create table data
        table_data = []
        for author in authors:
            short_name = self._get_short_name(author['full_name'])
            row = [
                short_name,
                author['time_spent'],
                author['equipment']
            ]
            table_data.append(row)
            print(f"  Row: {row}")

        # Headers
        headers = [
            "ФИО",
            "Время, затраченное на создание РИД, час",
            "Наименование и инвентарный номер оборудования, используемого для создания РИД и его место нахождения"
        ]

        # Find and replace placeholder
        success = self._replace_placeholder_with_table(doc, 'TABLE2', table_data, headers)

        if not success:
            print("  ⚠️ TABLE2 placeholder not found, adding table at the end")
            self._add_table_at_end(doc, headers, table_data, "Таблица 2")

    def _replace_placeholder_with_table(self, doc, placeholder, table_data, headers):
        """Replace placeholder with table"""
        placeholder_variants = [
            f'{{{placeholder}}}',
            f'{{{{{placeholder}}}}}',
            placeholder,
            f'[{placeholder}]'
        ]

        # Search in paragraphs
        for para_idx, paragraph in enumerate(doc.paragraphs):
            for variant in placeholder_variants:
                if variant in paragraph.text:
                    print(f"  Found '{variant}' placeholder in paragraph at index {para_idx}")

                    # Create table
                    table = self._create_formatted_table(headers, table_data)

                    # Replace paragraph with table
                    parent = paragraph._element.getparent()
                    parent.insert(parent.index(paragraph._element) + 1, table._element)
                    parent.remove(paragraph._element)

                    # Add space after table
                    new_para = OxmlElement('w:p')
                    parent.insert(parent.index(table._element) + 1, new_para)

                    return True

        print(f"  ❌ No '{placeholder}' placeholder found in document")
        return False

    def _add_table_at_end(self, doc, headers, data, title=None):
        """Add table at the end of document"""
        if title:
            p = doc.add_paragraph(title)
            p.runs[0].bold = True
            p.runs[0].font.name = 'Times New Roman'
            p.runs[0].font.size = Pt(12)

        table = self._create_formatted_table(headers, data)
        doc._body._element.append(table._element)

        # Add empty paragraph after table
        doc.add_paragraph()

    def _create_formatted_table(self, headers, data):
        """Create formatted table with headers and data"""
        # Create table
        table = Document().add_table(rows=len(data) + 1, cols=len(headers))

        # Remove default style
        try:
            table.style = None
        except:
            pass

        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set column widths
        if len(headers) == 3:
            widths = [Inches(2.5), Inches(3.5), Inches(1.5)]
        else:
            widths = [Inches(2.0), Inches(1.5), Inches(3.0)]

        for i, width in enumerate(widths):
            if i < len(table.columns):
                for cell in table.columns[i].cells:
                    cell.width = width

        # Apply borders
        self._apply_table_borders(table)

        # Fill header row
        header_row = table.rows[0]
        for i, header_text in enumerate(headers):
            cell = header_row.cells[i]
            cell.text = ''

            # Split by newlines for multi-line headers
            lines = header_text.split('\n')
            for line_idx, line in enumerate(lines):
                if line_idx > 0:
                    cell.add_paragraph()
                p = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(line.strip())
                run.font.name = 'Times New Roman'
                run.font.size = Pt(10)
                run.bold = True

        # Fill data rows
        for row_idx, row_data in enumerate(data):
            table_row = table.rows[row_idx + 1]
            for col_idx, cell_data in enumerate(row_data):
                cell = table_row.cells[col_idx]
                cell.text = str(cell_data) if cell_data is not None else ""

                # Apply formatting to all paragraphs in cell
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                    # Center align for percentage and time columns
                    if col_idx == 2 or (len(headers) == 3 and col_idx == 2):
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)

        return table

    def _create_signature_blocks(self, doc, authors):
        """Create signature blocks for each author - ИСПРАВЛЕННАЯ ВЕРСИЯ"""
        print(f"\n[_create_signature_blocks] Creating signatures for {len(authors)} authors")

        # Find placeholder
        placeholder_variants = ['{SIGNATURE}', '{{SIGNATURE}}', 'SIGNATURE', '[SIGNATURE]']

        # Ищем placeholder во всех параграфах
        for para_idx, paragraph in enumerate(doc.paragraphs):
            for variant in placeholder_variants:
                if variant in paragraph.text:
                    print(f"  Found '{variant}' placeholder at paragraph {para_idx}")

                    # Получаем родительский элемент и индекс
                    parent = paragraph._element.getparent()
                    para_index = parent.index(paragraph._element)

                    # Удаляем placeholder параграф
                    parent.remove(paragraph._element)

                    # УБИРАЕМ заголовок "Подписи сторон:" - его не должно быть
                    # para_index остаётся тем же

                    # Добавляем подписи для каждого автора
                    for author_idx, author in enumerate(authors):
                        # Строка 1: подчеркивания - 5.5 см и 3.25 см
                        line1_para = OxmlElement('w:p')
                        parent.insert(para_index, line1_para)
                        para_index += 1

                        line1_r = OxmlElement('w:r')
                        line1_t = OxmlElement('w:t')
                        # 5.5 см примерно 22 символа, 3.25 см примерно 13 символов
                        line1_t.text = "_" * 31 + "\t\t\t\t\t\t\t\t\t" + "_" * 17  # 3 табуляции между
                        line1_r.append(line1_t)
                        line1_para.append(line1_r)

                        # Форматируем подчеркивания
                        line1_rPr = OxmlElement('w:rPr')
                        line1_rFonts = OxmlElement('w:rFonts')
                        line1_rFonts.set(qn('w:ascii'), 'Times New Roman')
                        line1_rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                        line1_rPr.append(line1_rFonts)
                        line1_sz = OxmlElement('w:sz')
                        line1_sz.set(qn('w:val'), '20')
                        line1_rPr.append(line1_sz)
                        line1_r.insert(0, line1_rPr)

                        # Строка 2: метки курсивом с 3 табуляциями между
                        line2_para = OxmlElement('w:p')
                        parent.insert(para_index, line2_para)
                        para_index += 1

                        line2_r = OxmlElement('w:r')
                        line2_t = OxmlElement('w:t')
                        # (ФИО) и (подпись) с 3 табуляциями между
                        line2_t.text = "(ФИО)\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t(подпись)"  # 3 табуляции
                        line2_r.append(line2_t)
                        line2_para.append(line2_r)

                        # Форматируем метки (курсив, меньший размер)
                        line2_rPr = OxmlElement('w:rPr')
                        line2_i = OxmlElement('w:i')
                        line2_i.set(qn('w:val'), 'true')
                        line2_rPr.append(line2_i)
                        line2_rFonts = OxmlElement('w:rFonts')
                        line2_rFonts.set(qn('w:ascii'), 'Times New Roman')
                        line2_rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                        line2_rPr.append(line2_rFonts)
                        line2_sz = OxmlElement('w:sz')
                        line2_sz.set(qn('w:val'), '16')
                        line2_rPr.append(line2_sz)
                        line2_r.insert(0, line2_rPr)

                        # Добавляем отступ между подписями (кроме последней)
                        if author_idx < len(authors) - 1:
                            space_para = OxmlElement('w:p')
                            parent.insert(para_index, space_para)
                            para_index += 1

                    return True

        # If no placeholder found, add at the end
        print("  ⚠️ No SIGNATURE placeholder found, adding at the end")
        self._create_signatures_at_end(doc, authors)
        return False

    def _create_signatures_at_end(self, doc, authors):
        """Create signatures at the end of document - ИСПРАВЛЕННАЯ ВЕРСИЯ"""
        doc.add_paragraph()

        # УБИРАЕМ заголовок "Подписи сторон:" - его не должно быть

        # Add signatures
        for author_idx, author in enumerate(authors):
            # Line 1 - 5.5 см и 3.25 см с 3 табуляциями
            p1 = doc.add_paragraph()
            p1.paragraph_format.line_spacing = 1
            p1.paragraph_format.space_after = Pt(0)

            # 5.5 см примерно 22 символа, 3.25 см примерно 13 символов
            run1 = p1.add_run("_" * 31 + "\t\t\t\t\t\t\t\t\t" + "_" * 17)  # 3 табуляции между
            run1.font.name = 'Times New Roman'
            run1.font.size = Pt(10)

            # Line 2 - (ФИО) и (подпись) с 3 табуляциями
            p2 = doc.add_paragraph()
            p2.paragraph_format.line_spacing = 1
            p2.paragraph_format.space_after = Pt(12)

            run2 = p2.add_run(
                "(ФИО)\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t(подпись)")  # 3 табуляции
            run2.font.name = 'Times New Roman'
            run2.font.size = Pt(8)
            run2.italic = True  # Курсив

            # Add space between signatures
            if author_idx < len(authors) - 1:
                doc.add_paragraph()

    def _extract_product_name(self, application):
        """Extract product name from application"""
        if not application:
            return 'Тестовый продукт'

        if isinstance(application, tuple):
            print(f"  [_extract_product_name] Application is tuple with {len(application)} elements")

            # Индекс 8 - название продукта
            if len(application) > 8 and application[8]:
                value = str(application[8]).strip()
                if value and value.lower() not in ['admin', 'none'] and not value.isdigit():
                    print(f"    Found product name at index 8: '{value}'")
                    return value

            # Индекс 1 - название заявки (резервный вариант)
            if len(application) > 1 and application[1]:
                value = str(application[1]).strip()
                if value and value.lower() not in ['admin', 'none'] and not value.isdigit():
                    print(f"    Found product name at index 1: '{value}'")
                    return value

        print(f"  ⚠️ No product name found, using default")
        return 'Тестовый продукт'

    def _extract_nir_info(self, application):
        """Extract NIR information from application"""
        if not application:
            return "Научно-исследовательская работа"

        if isinstance(application, tuple):
            # Индекс 7 - NIR информация
            if len(application) > 7 and application[7]:
                value = str(application[7]).strip()
                if value and value.lower() not in ['admin', 'none']:
                    return value

        # Fallback
        product_name = self._extract_product_name(application)
        return f"Научно-исследовательская работа по теме: {product_name}"

    def _get_short_name(self, full_name):
        """Get short name (last name + initials)"""
        parts = full_name.strip().split()
        if len(parts) >= 2:
            surname = parts[0]
            name_initial = parts[1][0].upper() + '.'
            if len(parts) > 2:
                patronymic_initial = parts[2][0].upper() + '.'
                return f"{surname} {name_initial}{patronymic_initial}"
            else:
                return f"{surname} {name_initial}"
        return full_name

    def _to_genitive_department(self, department):
        """Convert department to genitive case"""
        if not department:
            return ''

        dept = department.strip()
        dept_lower = dept.lower()

        rules = [
            ('кафедра$', 'кафедры'),
            ('отдел$', 'отдела'),
            ('лаборатория$', 'лаборатории'),
            ('факультет$', 'факультета'),
            ('институт$', 'института'),
            ('центр$', 'центра'),
            ('управление$', 'управления'),
            ('служба$', 'службы'),
            ('бюро$', 'бюро'),
            ('сектор$', 'сектора'),
            ('группа$', 'группы'),
            ('команда$', 'команды'),
            ('ая$', 'ой'),
            ('яя$', 'ей'),
            ('ий$', 'его'),
            ('ый$', 'ого'),
            ('ой$', 'ого'),
            ('а$', 'ы'),
            ('я$', 'и'),
            ('о$', 'а'),
            ('е$', 'я'),
            ('ь$', 'я'),
        ]

        for pattern, replacement in rules:
            if re.search(pattern, dept_lower):
                result = re.sub(pattern[:-1], replacement, dept)
                return result

        return dept + 'а'

    def _apply_table_borders(self, table):
        """Apply borders to all cells in the table"""
        try:
            tbl = table._element
            tblPr = tbl.tblPr

            existing_borders = tblPr.findall(qn('w:tblBorders'))
            for border in existing_borders:
                tblPr.remove(border)

            tblBorders = OxmlElement('w:tblBorders')

            border_types = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
            for border_name in border_types:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), '4')
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), '000000')
                tblBorders.append(border)

            tblPr.append(tblBorders)

            for row in table.rows:
                for cell in row.cells:
                    tc = cell._element
                    tcPr = tc.get_or_add_tcPr()

                    existing_cell_borders = tcPr.findall(qn('w:tcBorders'))
                    for border in existing_cell_borders:
                        tcPr.remove(border)

                    tcBorders = OxmlElement('w:tcBorders')
                    for border_name in ['top', 'left', 'bottom', 'right']:
                        border = OxmlElement(f'w:{border_name}')
                        border.set(qn('w:val'), 'single')
                        border.set(qn('w:sz'), '4')
                        border.set(qn('w:space'), '0')
                        border.set(qn('w:color'), '000000')
                        tcBorders.append(border)
                    tcPr.append(tcBorders)

        except Exception as e:
            print(f"  ⚠️ Error applying table borders: {e}")

    def _replace_in_paragraphs(self, doc, replacements):
        """Replace text in paragraphs with underlining for PRODUCT_NAME and NIR"""
        for paragraph in doc.paragraphs:
            original_text = paragraph.text
            if original_text:
                # Сохраняем оригинальный текст для проверки
                temp_text = original_text

                # Проверяем, нужно ли подчеркивать (только для PRODUCT_NAME и NIR)
                need_underline = any(key in temp_text.upper() for key in ['PRODUCT_NAME', 'NIR'])

                # Заменяем текст - ТОЛЬКО полные совпадения
                new_text = original_text
                for old, new in replacements.items():
                    # Заменяем только если это точное совпадение (с учетом регистра)
                    if old in new_text:
                        new_text = new_text.replace(old, new)
                        print(f"  Replaced '{old}' with '{new}' in: '{original_text}'")

                if new_text != original_text:
                    print(f"  Replaced in paragraph: '{original_text}' -> '{new_text}' (underline: {need_underline})")

                    # Полностью очищаем параграф
                    paragraph.clear()

                    # Добавляем новый текст с нужным форматированием
                    run = paragraph.add_run(new_text)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

                    if need_underline:
                        run.underline = True

    def _replace_in_tables(self, doc, replacements):
        """Replace text in tables with underlining for PRODUCT_NAME and NIR"""
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        original_text = paragraph.text
                        if original_text:
                            # Проверяем, нужно ли подчеркивать
                            need_underline = any(key in original_text.upper() for key in ['PRODUCT_NAME', 'NIR'])

                            # Заменяем текст
                            new_text = original_text
                            for old, new in replacements.items():
                                if old in new_text:
                                    new_text = new_text.replace(old, new)

                            if new_text != original_text:
                                print(f"  Replaced in table cell: '{original_text}' -> '{new_text}'")

                                # Очищаем параграф
                                paragraph.clear()

                                # Добавляем новый текст с нужным форматированием
                                run = paragraph.add_run(new_text)
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(12)

                                if need_underline:
                                    run.underline = True
