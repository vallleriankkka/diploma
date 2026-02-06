import os
import io
from datetime import datetime, date
from decimal import Decimal
from docx import Document
from flask import current_app


class AssignmentAgreementGenerator:
    """Генератор договора об отчуждении исключительного права - ДЛЯ ВСЕХ АВТОРОВ"""

    def __init__(self, template_name='assignment_agreement_template.docx'):
        self.template_name = template_name
        self.template_path = self._get_template_path()
        self.setting_dao = None
        self.author_dao = None

    def _get_template_path(self):
        """Получение пути к шаблону"""
        return os.path.join(
            current_app.root_path,
            'templates',
            'documents',
            self.template_name
        )

    def set_daos(self, setting_dao, author_dao=None):
        """Установка DAO объектов для доступа к данным"""
        self.setting_dao = setting_dao
        self.author_dao = author_dao

    def generate_for_all_authors(self, application, authors_data):
        """Генерация договоров для ВСЕХ авторов с учетом процента вклада"""
        try:
            if not os.path.exists(self.template_path):
                print(f"Шаблон не найден: {self.template_path}")
                return []

            if not authors_data:
                print("Нет данных об авторах")
                return []

            # Получаем настройки
            settings_data = self._get_settings_data()
            if not settings_data:
                print("Не удалось получить настройки")
                return []

            print(f"DEBUG: Total authors received: {len(authors_data)}")

            # Получаем базовую сумму из приложения или настроек
            base_amount = self._get_base_amount(application, settings_data)
            print(f"DEBUG: Base amount: {base_amount}")

            # Генерируем для ВСЕХ авторов без проверки на сотрудников
            documents = []
            for author_index, author in enumerate(authors_data):
                print(f"\nDEBUG: Processing author #{author_index}")

                # Проверяем, есть ли данные автора
                if not author:
                    print(f"  Author #{author_index} - empty data, skipping")
                    continue

                print(f"  Author tuple length: {len(author)}")

                # Получаем процент вклада автора
                share_percentage = self._get_author_share_percentage(author, author_index)
                print(f"  Author #{author_index} - Share percentage: {share_percentage}%")

                # Рассчитываем сумму для данного автора
                author_amount = self._calculate_author_amount(base_amount, share_percentage)
                print(f"  Author #{author_index} - Calculated amount: {author_amount}")

                # Генерируем договор для этого автора
                doc_buffer = self._generate_single_author(
                    application, author, authors_data, settings_data, author_index, author_amount, share_percentage
                )
                if doc_buffer:
                    documents.append(doc_buffer)

            print(f"DEBUG: Generated {len(documents)} documents for {len(authors_data)} authors")
            return documents

        except Exception as e:
            print(f"Ошибка при генерации договоров: {e}")
            import traceback
            traceback.print_exc()
            return []

    def _generate_single_author(self, application, author, all_authors, settings_data, author_index, author_amount=None,
                                share_percentage=None):
        """Генерация договора для одного автора"""
        try:
            if not os.path.exists(self.template_path):
                print(f"Шаблон не найден: {self.template_path}")
                return None

            if not author:
                print("Нет данных об авторе")
                return None

            doc = Document(self.template_path)
            self._fill_template_for_single_author(
                doc, application, author, all_authors, settings_data,
                author_index, author_amount, share_percentage
            )

            # Сохраняем документ в буфер
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            return doc_buffer

        except Exception as e:
            print(f"Ошибка при генерации договора для автора #{author_index}: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _fill_template_for_single_author(self, doc, application, author, all_authors, settings_data, author_index,
                                         author_amount=None, share_percentage=None):
        """Заполнение шаблона данными для одного автора"""
        print(f"\n{'=' * 50}")
        print(f"DEBUG: FILL TEMPLATE FOR AUTHOR #{author_index}")
        print('=' * 50)

        # 1. Данные заявки (продукт)
        app_data = self._extract_application_data(application)
        print(f"DEBUG: Application data: {app_data}")

        # 2. Данные автора-отчуждателя
        author_data = self._extract_author_data(author, author_index)
        print(f"DEBUG: Author data (assignor): {author_data}")

        # 3. Добавляем процент вклада в данные автора
        if share_percentage is not None:
            author_data['share_percentage'] = f"{share_percentage}%"
            print(f"DEBUG: Author #{author_index} - Share percentage added: {share_percentage}%")

        # 4. Извлекаем конкретные данные из настроек
        assignee1_data = self._extract_assignee1_data(settings_data)
        print(f"DEBUG: Assignee1 data: {assignee1_data}")

        assignee2_data = self._extract_assignee2_data(settings_data)
        print(f"DEBUG: Assignee2 data: {assignee2_data}")

        power_of_attorney_data = self._extract_power_of_attorney_data(settings_data)
        print(f"DEBUG: Power of attorney data: {power_of_attorney_data}")

        # 5. Создаем данные для суммы с учетом процента вклада
        sum_data = {}
        if author_amount is not None:
            sum_data['value'] = author_amount
            sum_data['formatted'] = self._format_currency(author_amount)
            print(f"DEBUG: Using calculated author amount: {author_amount}")
        else:
            # Используем базовую сумму из настроек
            sum_data = self._extract_sum_data(settings_data)
            print(f"DEBUG: Using sum from settings: {sum_data}")

        # 6. Остальные авторы (SOA - Other Authors)
        other_authors = self._get_other_authors_formatted(all_authors, author, author_index)
        print(f"DEBUG: Other authors: '{other_authors}'")

        # 7. Форматируем POWER_OF_ATTORNEY
        power_of_attorney = self._format_power_of_attorney(power_of_attorney_data)
        print(f"DEBUG: Formatted POWER_OF_ATTORNEY: '{power_of_attorney}'")

        # 8. Форматируем SUM с учетом процента вклада
        sum_amount = self._format_sum_with_share(sum_data, share_percentage)
        print(f"DEBUG: Formatted SUM with share: '{sum_amount}'")

        # Словарь замен
        replacements = {
            # ФИО автора-отчуждателя
            'FULLNAME': author_data.get('full_name', 'Автор не указан'),

            # Должность
            'POSITION': author_data.get('position', ''),

            # Должность в родительном падеже
            'POSGEN': author_data.get('position_genitive', ''),

            # Подразделение
            'DEPARTMENT': author_data.get('department', ''),

            # Подразделение в родительном падеже
            'DEPGEN': author_data.get('department_genitive', ''),

            # СНИЛС
            'SNILS': author_data.get('snils', ''),

            # ИНН
            'INN': author_data.get('inn', ''),

            # Дата рождения компоненты
            'DAY': author_data.get('birth_day', ''),

            'MONTH': author_data.get('birth_month', ''),

            'YEAR': author_data.get('birth_year', ''),

            # Паспортные данные
            'SERIES': author_data.get('passport_series', ''),

            'NUMBER': author_data.get('passport_number', ''),

            'DATEP': author_data.get('passport_date', ''),

            'BY': author_data.get('passport_by', ''),

            'CODE': author_data.get('passport_code', ''),

            # Адрес
            'ADDRESS': author_data.get('address', 'Адрес не указан'),

            # Продукт/программа
            'PRODUCTNAME': app_data.get('product_name', 'Не указано'),

            # Доверенность (из настроек)
            'POWER': power_of_attorney,

            # Остальные авторы в скобках
            'SOA': other_authors,

            # Сумма договора (с учетом процента вклада)
            'SUM': sum_amount,

            # Сторона-приобретатель 1 (в родительном падеже)
            'POSTA': assignee1_data.get('post_genitive', ''),

            'ASSIGNEE1': assignee1_data.get('name_genitive', ''),

            # Сторона-приобретатель 2 (из настроек)
            '{{POST2}}': assignee2_data.get('post', ''),
            '{POST2}': assignee2_data.get('post', ''),
            'POST2': assignee2_data.get('post', ''),

            '{{ASSIGNEE2}}': assignee2_data.get('name', ''),
            '{ASSIGNEE2}': assignee2_data.get('name', ''),
            'ASSIGNEE2': assignee2_data.get('name', ''),

            # Текст с процентом вклада
            '{{SHARE_TEXT}}': f"({share_percentage or '___'}% вклада)" if share_percentage else "",
            '{SHARE_TEXT}': f"({share_percentage or '___'}% вклада)" if share_percentage else "",
            'SHARE_TEXT': f"({share_percentage or '___'}% вклада)" if share_percentage else "",
        }

        # Отладочный вывод
        print(f"\nDEBUG: Key replacements for author #{author_index}:")
        print(f"  FULLNAME: '{replacements['{{FULLNAME}}']}'")
        print(f"  SHARE_PERCENTAGE: '{replacements['{{SHARE_PERCENTAGE}}']}'")
        print(f"  SHARE_TEXT: '{replacements['{{SHARE_TEXT}}']}'")
        print(f"  SUM: '{replacements['{{SUM}}']}'")
        print(f"  PRODUCTNAME: '{replacements['{{PRODUCTNAME}}']}'")

        print(f"\n{'=' * 50}")
        print(f"DEBUG: END FILL TEMPLATE FOR AUTHOR #{author_index}")
        print('=' * 50)

        self._replace_in_paragraphs_preserve_formatting(doc, replacements)
        self._replace_in_tables_preserve_formatting(doc, replacements)

    def _get_base_amount(self, application, settings_data):
        """Получение базовой суммы из приложения или настроек"""
        try:
            # Если в приложении не найдено, пробуем из настроек
            if settings_data and len(settings_data) > 9 and settings_data[9]:
                try:
                    amount = float(settings_data[9])
                    print(f"DEBUG: Using amount from settings: {amount}")
                    return amount
                except:
                    pass

            # Значение по умолчанию
            print("DEBUG: Using default amount 0.00")
            return 0.00

        except Exception as e:
            print(f"DEBUG: Error getting base amount: {e}")
            return 0.00

    def _get_author_share_percentage(self, author, author_index):
        """Получение процента вклада автора из данных автора"""
        try:
            if not author or not isinstance(author, tuple):
                return None

            # Процент вклада может быть в разных полях
            # Обычно в структуре application_author процент может быть в поле 4 или 5
            possible_share_fields = [3]  # Индексы, где может храниться процент

            for field_index in possible_share_fields:
                if len(author) > field_index and author[field_index] is not None:
                    share_value = author[field_index]

                    if isinstance(share_value, (int, float, Decimal)):
                        percentage = float(share_value)
                        if 0 <= percentage <= 100:
                            print(f"DEBUG: Author #{author_index} - Found share {percentage}% at field {field_index}")
                            return percentage

                    elif isinstance(share_value, str):
                        try:
                            # Очищаем строку и преобразуем в число
                            cleaned = ''.join(c for c in share_value if c.isdigit() or c in '.,')
                            if cleaned:
                                cleaned = cleaned.replace(',', '.')
                                percentage = float(cleaned)
                                if 0 <= percentage <= 100:
                                    print(
                                        f"DEBUG: Author #{author_index} - Found share string {percentage}% at field {field_index}")
                                    return percentage
                        except:
                            pass

            # Если процент не найден, возвращаем None
            print(f"DEBUG: Author #{author_index} - No share percentage found")
            return None

        except Exception as e:
            print(f"DEBUG: Error getting author share percentage: {e}")
            return None

    def _calculate_author_amount(self, base_amount, share_percentage):
        """Расчет суммы для автора на основе процента вклада"""
        try:
            if share_percentage is None:
                # Если процент не указан, возвращаем полную сумму
                return base_amount

            # Рассчитываем сумму: базовая сумма * процент / 100
            author_amount = (base_amount * share_percentage) / 100.0

            # Округляем до 2 знаков после запятой
            author_amount = round(author_amount, 2)

            print(f"DEBUG: Calculated author amount: {base_amount} * {share_percentage}% = {author_amount}")
            return author_amount

        except Exception as e:
            print(f"DEBUG: Error calculating author amount: {e}")
            return base_amount

    def _format_sum_with_share(self, sum_data, share_percentage=None):
        """Форматирование суммы договора с учетом процента вклада"""
        if not sum_data or not sum_data.get('value'):
            return "___________ (___________________) рублей"

        try:
            # Получаем значение суммы
            value_str = str(sum_data['value']).strip()

            # Пробуем преобразовать в целое число (только рубли)
            if '.' in value_str or ',' in value_str:
                value_clean = value_str.replace(',', '.')
                num_value = float(value_clean)
                # Округляем до целых рублей
                rubles = int(round(num_value))
            else:
                rubles = int(value_str)

            # Форматируем числовую часть
            num_part = f"{rubles:,}".replace(',', ' ')

            # Преобразуем рубли в пропись
            sum_in_words = self._number_to_words(rubles)

            # Добавляем форму слова "рубль"
            if 10 <= rubles % 100 <= 20:
                rub_form = 'рублей'
            elif rubles % 10 == 1:
                rub_form = 'рубль'
            elif 2 <= rubles % 10 <= 4:
                rub_form = 'рубля'
            else:
                rub_form = 'рублей'

            # Формируем базовую строку
            result = f"{num_part} ({sum_in_words} {rub_form}) рублей"

            # Добавляем информацию о проценте вклада, если он указан
            if share_percentage is not None:
                result += f" ({share_percentage}% вклада)"

            return result

        except Exception as e:
            print(f"DEBUG: Error formatting sum with share: {e}")
            return "___________ (___________________) рублей"

    def _get_settings_data(self):
        """Получение данных настроек из базы"""
        try:
            if self.setting_dao:
                # Получаем текущие настройки
                settings = self.setting_dao.get_current()

                if settings:
                    print(f"DEBUG: Got settings from DAO: {settings}")
                    return settings
                else:
                    print("DEBUG: No settings found in database")

            return None

        except Exception as e:
            print(f"DEBUG: Error getting settings: {e}")
            return None

    def _extract_assignee1_data(self, settings_data):
        """Извлечение данных стороны-приобретателя 1 из настроек"""
        result = {
            'post': 'Проректор по науке и инновациям',
            'post_genitive': 'Проректора по науке и инновациям',
            'name': 'А.В. Башкиров',
            'name_genitive': 'А.В. Башкирова',
        }

        if not settings_data:
            print("DEBUG: No settings data for assignee1")
            return result

        try:
            # Имя проректора
            if len(settings_data) > 5 and settings_data[5]:
                result['name'] = str(settings_data[5]).strip()
                result['name_genitive'] = self._to_genitive_name_full(result['name'])
                print(f"DEBUG: Found assignee1 name at index 5: '{result['name']}'")

            # Должность проректора - индекс 6, а не 9!
            if len(settings_data) > 6 and settings_data[6]:
                result['post'] = str(settings_data[6]).strip()
                result['post_genitive'] = self._to_genitive_position(result['post'])
                print(f"DEBUG: Found assignee1 post at index 6: '{result['post']}'")

        except Exception as e:
            print(f"DEBUG: Error extracting assignee1 data: {e}")

        return result

    def _to_genitive_name_full(self, name):
        """Полное преобразование имени в родительный падеж (все части)"""
        if not name:
            return ''

        name = name.strip()
        parts = name.split()

        if len(parts) < 2:
            return self._to_genitive_name(name)

        # Преобразуем каждую часть
        genitive_parts = []
        for i, part in enumerate(parts):
            part_lower = part.lower()

            if i == 0:  # Фамилия
                if part_lower.endswith('ов'):
                    gen_part = part + 'а'
                elif part_lower.endswith('ев'):
                    gen_part = part + 'а'
                elif part_lower.endswith('ин'):
                    gen_part = part + 'а'
                elif part_lower.endswith('ский'):
                    gen_part = part[:-2] + 'ого'
                elif part_lower.endswith('цкий'):
                    gen_part = part[:-2] + 'ого'
                elif part_lower.endswith('ой'):
                    gen_part = part[:-2] + 'ого'
                elif part_lower.endswith('ий'):
                    gen_part = part[:-2] + 'ого'
                elif part_lower.endswith('ый'):
                    gen_part = part[:-2] + 'ого'
                else:
                    gen_part = part + 'а'

            else:  # Имя и отчество
                if part_lower.endswith('ей'):
                    if part_lower == 'алексей':
                        gen_part = part[:-2] + 'ея'
                    elif part_lower == 'николай':
                        gen_part = part[:-2] + 'ая'
                    elif part_lower == 'сергей':
                        gen_part = part[:-2] + 'ея'
                    else:
                        gen_part = part[:-2] + 'я'
                elif part_lower.endswith('ий'):
                    if part_lower == 'дмитрий':
                        gen_part = part[:-2] + 'ия'
                    elif part_lower == 'викторий':
                        gen_part = part[:-2] + 'ия'
                    else:
                        gen_part = part[:-2] + 'ого'
                elif part_lower.endswith('я'):
                    gen_part = part[:-1] + 'и'
                elif part_lower.endswith('а'):
                    gen_part = part[:-1] + 'ы'
                elif part_lower.endswith('ь'):
                    gen_part = part[:-1] + 'я'
                else:
                    gen_part = part + 'а'

            # Сохраняем заглавную букву
            if part and part[0].isupper() and gen_part:
                gen_part = gen_part[0].upper() + gen_part[1:]

            genitive_parts.append(gen_part)

        return ' '.join(genitive_parts)

    def _extract_assignee2_data(self, settings_data):
        """Извлечение данных стороны-приобретателя 2 из настроек"""
        result = {
            'post': 'Ректор',
            'name': 'И.И. Иванов',
        }

        if not settings_data:
            print("DEBUG: No settings data for assignee2")
            return result

        try:
            if len(settings_data) > 10 and settings_data[10]:
                result['name'] = str(settings_data[10]).strip()
                print(f"DEBUG: Found assignee2 name at index 10: '{result['name']}'")

            if len(settings_data) > 11 and settings_data[11]:
                result['post'] = str(settings_data[11]).strip()
                print(f"DEBUG: Found assignee2 post at index 11: '{result['post']}'")

        except Exception as e:
            print(f"DEBUG: Error extracting assignee2 data: {e}")

        return result

    def _extract_power_of_attorney_data(self, settings_data):
        """Извлечение данных доверенности из настроек"""
        result = {
            'number': '',
            'date': '',
            'date_formatted': '',
        }

        if not settings_data:
            print("DEBUG: No settings data for power of attorney")
            return result

        try:
            # Номер доверенности - обычно поле 7 в настройках
            if len(settings_data) > 7 and settings_data[7]:
                result['number'] = str(settings_data[7]).strip()
                print(f"DEBUG: Found power of attorney number: '{result['number']}'")

            # Дата доверенности - обычно поле 8 в настройках
            if len(settings_data) > 8 and settings_data[8]:
                date_value = settings_data[8]
                if isinstance(date_value, (datetime, date)):
                    result['date'] = date_value
                    result['date_formatted'] = date_value.strftime('%d.%m.%Y')
                    print(f"DEBUG: Found power of attorney date (datetime): {result['date_formatted']}")
                elif isinstance(date_value, str):
                    result['date'] = date_value
                    result['date_formatted'] = self._parse_date(date_value)
                    print(f"DEBUG: Found power of attorney date (string): {result['date_formatted']}")

        except Exception as e:
            print(f"DEBUG: Error extracting power of attorney data: {e}")

        return result

    def _extract_sum_data(self, settings_data):
        """Извлечение данных суммы из настроек"""
        result = {
            'value': '',
            'formatted': '',
        }

        if not settings_data:
            return result

        try:
            if len(settings_data) > 9 and settings_data[9]:
                sum_value = settings_data[9]
                result['value'] = str(sum_value)
                result['formatted'] = self._format_currency(sum_value)
                print(f"DEBUG: Found sum at index 9: '{result['value']}'")

        except Exception as e:
            print(f"DEBUG: Error extracting sum data: {e}")

        return result

    def _format_power_of_attorney(self, power_of_attorney_data):
        """Форматирование доверенности: только номер и дата"""
        if not power_of_attorney_data:
            return ""

        number = power_of_attorney_data.get('number', '').strip()
        date_formatted = power_of_attorney_data.get('date_formatted', '').strip()

        # Формируем строку "№ [номер] от [дата]"
        result_parts = []

        if number:
            # Добавляем знак № если его еще нет
            if not number.startswith('№'):
                result_parts.append(f"№ {number}")
            else:
                result_parts.append(number)

        if date_formatted:
            result_parts.append(f"от {date_formatted}")

        if result_parts:
            return " ".join(result_parts)

        return ""

    def _number_to_words(self, num):
        """Преобразование числа в пропись на русском языке"""
        # Базовые словари
        units_female = ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
        units_male = ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']
        teens = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать',
                 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать']
        tens = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят',
                'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто']
        hundreds = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот',
                    'шестьсот', 'семьсот', 'восемьсот', 'девятьсот']

        def convert_three_digits(n, gender='male'):
            """Конвертация трехзначного числа"""
            result = []

            # Сотни
            h = n // 100
            if h > 0:
                result.append(hundreds[h])

            # Десятки и единицы
            remainder = n % 100
            if remainder >= 20:
                t = remainder // 10
                u = remainder % 10
                result.append(tens[t])
                if u > 0:
                    if gender == 'female':
                        result.append(units_female[u])
                    else:
                        result.append(units_male[u])
            elif remainder >= 10:
                result.append(teens[remainder - 10])
            elif remainder > 0:
                if gender == 'female':
                    result.append(units_female[remainder])
                else:
                    result.append(units_male[remainder])

            return result

        if num == 0:
            return 'ноль'

        result = []

        # Миллиарды
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

        # Миллионы
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

        # Тысячи (женский род)
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

        # Единицы (мужской род)
        units = num % 1000
        if units > 0 or num == 0:
            result.extend(convert_three_digits(units))

        # Убираем пустые строки и соединяем
        words = [word for word in result if word]
        return ' '.join(words)

    def _extract_author_data(self, author, author_index=0):
        """Извлечение данных одного автора из кортежа (структура application_author)"""
        result = {
            'full_name': 'Автор не указан',
            'position': '',
            'department': '',
            'position_genitive': '',
            'department_genitive': '',
            'share_percentage': '',
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
            print(f"DEBUG: Author #{author_index} - Invalid author data")
            return result

        print(f"\nDEBUG: Processing author #{author_index} data")
        print(f"DEBUG: Author tuple length: {len(author)}")

        try:
            # 1. ФИО из индексов 7, 8, 9
            if len(author) > 9:
                surname = str(author[7]).strip() if author[7] else ''
                name = str(author[8]).strip() if author[8] else ''
                patronymic = str(author[9]).strip() if author[9] else ''
                result['full_name'] = f"{surname} {name} {patronymic}".strip()
                print(f"DEBUG: Author #{author_index} - ФИО: '{result['full_name']}'")

            # 2. Дата рождения из индекса 10
            if len(author) > 10 and author[10]:
                birth_date = author[10]
                if isinstance(birth_date, (datetime, date)):
                    birth_date_str = birth_date.strftime('%d.%m.%Y')
                else:
                    birth_date_str = str(birth_date)

                if birth_date_str:
                    day, month_name, month_num, year = self._parse_birth_date(birth_date_str)
                    result['birth_day'] = day
                    result['birth_month'] = month_name
                    result['birth_year'] = year
                    print(f"DEBUG: Author #{author_index} - Birth date: {day}.{month_name}.{year}")

            # 3. Адрес из индекса 12
            if len(author) > 12 and author[12] and isinstance(author[12], str):
                address = str(author[12]).strip()
                if address:
                    result['address'] = address
                    print(f"DEBUG: Author #{author_index} - Address: '{address[:50]}...'")

            # 4. Должность из индекса 14
            if len(author) > 14 and author[14] and isinstance(author[14], str):
                position = str(author[14]).strip()
                if position:
                    result['position'] = position
                    result['position_genitive'] = self._to_genitive_position(position)
                    print(f"DEBUG: Author #{author_index} - Position: '{position}'")

            # 5. Подразделение из индекса 18
            if len(author) > 18 and author[18] and isinstance(author[18], str):
                department = str(author[18]).strip()
                if department:
                    result['department'] = department
                    result['department_genitive'] = self._to_genitive_department(department)
                    print(f"DEBUG: Author #{author_index} - Department: '{department}'")

            # 6. Если есть author_id (индекс 1) и доступен author_dao, получаем полные данные
            if len(author) > 2 and author[2] and self.author_dao:
                try:
                    author_id = int(author[2])
                    print(f"DEBUG: Author #{author_index} - Getting full data from author_dao for id={author_id}")
                    author_details = self.author_dao.get_by_id(author_id)

                    if author_details:
                        print(f"DEBUG: Author #{author_index} - Got author details, length: {len(author_details)}")

                        # Паспортные данные
                        if len(author_details) > 6 and author_details[6] and isinstance(author_details[6], str):
                            passport_str = str(author_details[6]).strip()
                            passport_clean = ''.join(filter(str.isdigit, passport_str))

                            if len(passport_clean) >= 10:
                                series = passport_clean[:4]
                                result['passport_series'] = f"{series[:2]} {series[2:4]}"
                                result['passport_number'] = passport_clean[4:10]
                                print(f"DEBUG: Author #{author_index} - Passport from author_dao: '{passport_str}'")

                        # СНИЛС
                        if len(author_details) > 7 and author_details[7] and isinstance(author_details[7], str):
                            snils = str(author_details[7]).strip()
                            digits = ''.join(filter(str.isdigit, snils))
                            if len(digits) == 11:
                                result['snils'] = f"{digits[:3]}-{digits[3:6]}-{digits[6:9]} {digits[9:]}"
                            else:
                                result['snils'] = snils
                            print(f"DEBUG: Author #{author_index} - SNILS from author_dao: '{result['snils']}'")

                        # ИНН
                        if len(author_details) > 8 and author_details[8] and isinstance(author_details[8], str):
                            result['inn'] = str(author_details[8]).strip()
                            print(f"DEBUG: Author #{author_index} - INN from author_dao: '{result['inn']}'")

                        # Дата выдачи паспорта
                        if len(author_details) > 15 and author_details[15]:
                            issue_date = author_details[15]
                            if isinstance(issue_date, (datetime, date)):
                                result['passport_date'] = issue_date.strftime('%d.%m.%Y')
                            elif isinstance(issue_date, str) and issue_date.strip():
                                result['passport_date'] = self._parse_date(issue_date.strip())

                        # Кем выдан паспорт
                        if len(author_details) > 16 and author_details[16] and isinstance(author_details[16], str):
                            result['passport_by'] = str(author_details[16]).strip()

                        # Код подразделения
                        if len(author_details) > 17 and author_details[17] and isinstance(author_details[17], str):
                            code = str(author_details[17]).strip()
                            digits = ''.join(filter(str.isdigit, code))
                            if len(digits) == 6:
                                result['passport_code'] = f"{digits[:3]}-{digits[3:]}"
                            else:
                                result['passport_code'] = code
                except Exception as e:
                    print(f"DEBUG: Author #{author_index} - Error getting data from author_dao: {e}")

            # 7. Альтернативные источники данных (если нет author_dao)
            else:
                print(f"DEBUG: Author #{author_index} - No author_dao or author_id, using basic data")

                # Ищем СНИЛС в полях 5, 6, 11
                for i in [5, 6, 11]:
                    if len(author) > i and author[i] and isinstance(author[i], str):
                        val = str(author[i]).strip()
                        digits = ''.join(filter(str.isdigit, val))
                        if len(digits) == 11 and not result['snils']:
                            result['snils'] = f"{digits[:3]}-{digits[3:6]}-{digits[6:9]} {digits[9:]}"
                            print(f"DEBUG: Author #{author_index} - SNILS found at field {i}: '{result['snils']}'")
                            break
                        elif '-' in val and val.count('-') >= 2 and not result['snils']:
                            result['snils'] = val
                            print(f"DEBUG: Author #{author_index} - SNILS found at field {i}: '{result['snils']}'")
                            break

                # Ищем ИНН
                for i in [5, 6, 11]:
                    if len(author) > i and author[i] and isinstance(author[i], str):
                        val = str(author[i]).strip()
                        digits = ''.join(filter(str.isdigit, val))
                        if len(digits) in [10, 12] and not result['inn']:
                            result['inn'] = digits
                            print(f"DEBUG: Author #{author_index} - INN found at field {i}: '{result['inn']}'")
                            break

        except Exception as e:
            print(f"DEBUG: Error extracting author data: {e}")
            import traceback
            traceback.print_exc()

        # Заполняем заглушками, если данные не найдены
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
        result = {
            'product_name': 'Не указано',
        }

        if not application:
            return result

        try:
            if isinstance(application, dict):
                result['product_name'] = application.get('product_name', 'Не указано')
            elif isinstance(application, tuple):
                if len(application) > 8 and application[8]:
                    result['product_name'] = str(application[8]).strip()
                elif len(application) > 3 and application[3]:
                    result['product_name'] = str(application[3]).strip()
        except:
            pass

        return result

    def _to_genitive_name(self, name):
        """Преобразование полного имени в родительный падеж"""
        if not name:
            return ''

        name = name.strip()
        parts = name.split()

        if len(parts) == 0:
            return name

        # Преобразуем только фамилию (первую часть)
        last_name = parts[0]
        last_name_lower = last_name.lower()

        # Правила для русских фамилий
        if last_name_lower.endswith('ов'):
            last_name_gen = last_name + 'а'
        elif last_name_lower.endswith('ев'):
            last_name_gen = last_name + 'а'
        elif last_name_lower.endswith('ин'):
            last_name_gen = last_name + 'а'
        elif last_name_lower.endswith('ский'):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('цкий'):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('ой'):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('ий'):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('ый'):
            last_name_gen = last_name[:-2] + 'ого'
        elif last_name_lower.endswith('ая'):
            last_name_gen = last_name[:-2] + 'ой'
        elif last_name_lower.endswith('яя'):
            last_name_gen = last_name[:-2] + 'ей'
        else:
            last_name_gen = last_name + 'а'

        # Сохраняем заглавную букву
        if last_name[0].isupper() and last_name_gen:
            last_name_gen = last_name_gen[0].upper() + last_name_gen[1:]

        # Собираем обратно с остальными частями имени
        if len(parts) == 1:
            return last_name_gen
        else:
            return ' '.join([last_name_gen] + parts[1:])

    def _get_other_authors_formatted(self, all_authors, current_author, current_index):
        """Получение списка остальных авторов в формате 'в соавторстве (ФИО1, ФИО2, ...)'"""
        if not all_authors or len(all_authors) <= 1:
            return ""

        other_authors = []
        for i, author in enumerate(all_authors):
            if author == current_author or i == current_index:
                continue

            # Извлекаем ФИО из правильных индексов (7, 8, 9)
            full_name = None
            if len(author) > 9:
                surname = str(author[7]).strip() if author[7] else ''
                name = str(author[8]).strip() if author[8] else ''
                patronymic = str(author[9]).strip() if author[9] else ''
                full_name = f"{surname} {name} {patronymic}".strip()

            if full_name and full_name != '  ':
                other_authors.append(full_name)

        if other_authors:
            # Добавляем "в соавторстве" перед перечислением авторов
            return f"(в соавторстве {', '.join(other_authors)})"
        return ""

    def _parse_birth_date(self, birth_date_str):
        """Парсинг даты рождения на компоненты"""
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
                    if not day:
                        day = ""

                    month_num = parts[1]
                    if month_num:
                        month_int = int(month_num)
                        month_name = self._get_month_name_ru(month_int)

                    year = parts[2]

            elif '-' in birth_date_str:
                parts = birth_date_str.split('-')
                if len(parts) >= 3:
                    year = parts[0]

                    month_num = parts[1]
                    if month_num:
                        month_int = int(month_num)
                        month_name = self._get_month_name_ru(month_int)

                    day = parts[2].lstrip('0')
                    if not day:
                        day = ""

        except Exception as e:
            print(f"DEBUG: Error parsing birth date '{birth_date_str}': {e}")

        return day, month_name, month_num, year

    def _to_genitive_position(self, position):
        """Преобразование должности в родительный падеж"""
        if not position:
            return ''

        position_lower = position.lower().strip()
        position = position.strip()

        if position_lower == 'проректор по науке и инновациям':
            return 'Проректора по науке и инновации'
        elif position_lower == 'ректор':
            return 'Ректора'
        elif position_lower == 'директор':
            return 'Директора'
        elif position_lower == 'заведующий кафедрой':
            return 'Заведующего кафедрой'
        elif position_lower == 'начальник отдела':
            return 'Начальника отдела'
        elif position_lower == 'руководитель':
            return 'Руководителя'

        words = position.split()
        if len(words) == 0:
            return position

        first_word = words[0]
        first_word_lower = first_word.lower()

        if first_word_lower.endswith('ор'):
            first_word_gen = first_word + 'а'
        elif first_word_lower.endswith('ист'):
            first_word_gen = first_word + 'а'
        elif first_word_lower.endswith('ер'):
            first_word_gen = first_word + 'а'
        elif first_word_lower.endswith('ль'):
            if first_word_lower.endswith('ель'):
                first_word_gen = first_word[:-2] + 'ля'
            else:
                first_word_gen = first_word[:-2] + 'ля'
        elif first_word_lower.endswith('ой'):
            first_word_gen = first_word[:-2] + 'ого'
        elif first_word_lower.endswith('ий'):
            first_word_gen = first_word[:-2] + 'ого'
        elif first_word_lower.endswith('ый'):
            first_word_gen = first_word[:-2] + 'ого'
        elif first_word_lower.endswith('а'):
            first_word_gen = first_word[:-1] + 'ы'
        elif first_word_lower.endswith('я'):
            first_word_gen = first_word[:-1] + 'и'
        else:
            first_word_gen = first_word + 'а'

        if len(words) == 1:
            return first_word_gen

        if first_word[0].isupper() and first_word_gen:
            first_word_gen = first_word_gen[0].upper() + first_word_gen[1:] if first_word_gen else first_word_gen

        return ' '.join([first_word_gen] + words[1:])

    def _to_genitive_department(self, department):
        """Преобразование подразделения в родительный падеж"""
        if not department:
            return ''

        dept = department.strip()
        dept_lower = dept.lower()

        if 'кафедр' in dept_lower:
            if dept_lower.endswith('кафедра'):
                return dept[:-6] + 'кафедры'
            elif dept_lower.endswith('кафедры'):
                return dept
            elif dept_lower.endswith('кафедре'):
                return dept[:-5] + 'ы'
            elif dept_lower.endswith('кафедрой'):
                return dept[:-7] + 'ы'
            else:
                return dept + 'ы'

        elif 'отдел' in dept_lower:
            if dept_lower.endswith('отдел'):
                return dept[:-5] + 'отдела'
            elif dept_lower.endswith('отдела'):
                return dept
            elif dept_lower.endswith('отделе'):
                return dept[:-5] + 'а'
            elif dept_lower.endswith('отделом'):
                return dept[:-6] + 'а'
            else:
                return dept + 'а'

        elif 'лаборатор' in dept_lower:
            if dept_lower.endswith('лаборатория'):
                return dept[:-11] + 'лаборатории'
            elif dept_lower.endswith('лаборатории'):
                return dept
            elif dept_lower.endswith('лабораторией'):
                return dept[:-12] + 'ии'
            else:
                return dept + 'ии'

        elif 'факультет' in dept_lower:
            if dept_lower.endswith('факультет'):
                return dept[:-9] + 'факультета'
            elif dept_lower.endswith('факультета'):
                return dept
            elif dept_lower.endswith('факультетом'):
                return dept[:-11] + 'а'
            else:
                return dept + 'а'

        elif 'институт' in dept_lower:
            if dept_lower.endswith('институт'):
                return dept[:-8] + 'института'
            elif dept_lower.endswith('института'):
                return dept
            elif dept_lower.endswith('институтом'):
                return dept[:-10] + 'а'
            else:
                return dept + 'а'

        elif 'центр' in dept_lower:
            if dept_lower.endswith('центр'):
                return dept[:-5] + 'центра'
            elif dept_lower.endswith('центра'):
                return dept
            elif dept_lower.endswith('центром'):
                return dept[:-6] + 'а'
            else:
                return dept + 'а'

        if dept_lower.endswith('ая'):
            return dept[:-2] + 'ой'
        elif dept_lower.endswith('яя'):
            return dept[:-2] + 'ей'
        elif dept_lower.endswith('ий'):
            return dept[:-2] + 'его'
        elif dept_lower.endswith('ый'):
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

    def _format_date(self, date_value):
        """Форматирование даты"""
        try:
            if isinstance(date_value, (datetime, date)):
                return date_value.strftime('%d.%m.%Y')
            elif isinstance(date_value, str):
                return self._parse_date(date_value)
            return str(date_value)
        except:
            return str(date_value) if date_value else ''

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
                cleaned = ''.join(c for c in value if c.isdigit() or c == '.' or c == ',')
                cleaned = cleaned.replace(',', '.')
                if cleaned:
                    num_value = float(cleaned)
                else:
                    return str(value)
            else:
                return str(value)

            if num_value.is_integer():
                return f"{int(num_value):,}".replace(',', ' ') + ' рублей'
            else:
                return f"{num_value:,.2f}".replace(',', ' ') + ' рублей'

        except Exception as e:
            print(f"DEBUG: Error formatting currency '{value}': {e}")
            return str(value)

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
