import os
from datetime import datetime, date
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from flask import current_app


class ConclusionGenerator:
    """Document generator for conclusion about open publication"""

    def __init__(self, template_name='conclusion_open_publication.docx'):
        self.template_name = template_name
        self.template_path = self._get_template_path()

    def _get_template_path(self):
        """Get template path"""
        return os.path.join(
            current_app.root_path,
            'templates',
            'documents',
            self.template_name
        )

    def generate(self, application, commission_data=None, settings_data=None, authors_data=None):
        """Generate conclusion document"""
        try:
            if not os.path.exists(self.template_path):
                print(f"ERROR: Template not found: {self.template_path}")
                return None

            doc = Document(self.template_path)
            self._fill_template(doc, application, commission_data, settings_data, authors_data)
            return doc

        except Exception as e:
            print(f"ERROR in generate: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _fill_template(self, doc, application, commission_data, settings_data, authors_data=None):
        """Fill template with data"""
        print("\n" + "=" * 80)
        print("DEBUG: START Filling conclusion template...")
        print("=" * 80)

        # 1. Print application data structure
        print("\n=== APPLICATION DATA ===")
        if application:
            print(f"Type: {type(application)}, Length: {len(application)}")
            for i, field in enumerate(application):
                field_type = type(field).__name__
                field_str = str(field)[:100] if field else "None/Empty"
                print(f"[{i}] {field_type}: {field_str}")
        else:
            print("Application data is None")

        # 2. Get product/program name
        product_name = self._extract_product_name(application)
        print(f"\n=== PRODUCT NAME ===")
        print(f"Extracted: '{product_name}'")

        # 3. Get authors data
        print("\n=== AUTHORS DATA ===")
        if authors_data is None:
            print("No authors_data passed, extracting from application")
            authors_data = self._extract_authors_data(application)
        else:
            print(
                f"Authors data passed directly, type: {type(authors_data)}, length: {len(authors_data) if hasattr(authors_data, '__len__') else 'N/A'}")
            authors_data = self._extract_authors_from_list(authors_data)

        print(f"Found {len(authors_data['full_names_list']) if authors_data else 0} authors")
        print(f"Authors list: {authors_data.get('full_names_list', [])}")
        print(f"Authors formatted: {authors_data.get('authors_list_formatted', '')}")
        print(f"Authors genitive: {authors_data.get('authors_genitive_formatted', '')}")

        # 4. Get commission data (experts)
        print("\n=== COMMISSION DATA ===")
        if commission_data:
            print(f"Commission data type: {type(commission_data)}, length: {len(commission_data)}")
            if commission_data and len(commission_data) > 0:
                print(f"First expert tuple: {commission_data[0]}")
                print(f"First expert length: {len(commission_data[0])}")
                for i, field in enumerate(commission_data[0]):
                    print(f"  Expert[0][{i}]: {field}")
        else:
            print("No commission data provided")

        commission_info = self._extract_commission_data(commission_data) if commission_data else {}
        print(f"Commission info keys: {list(commission_info.keys())}")

        # Получаем специальные форматированные поля
        memchair = self._get_chairman1(commission_info)  # Только ФИО председателя
        members = self._get_expert2(commission_info)  # Все члены комиссии с ролями (ВМЕСТО CHAIRMAN2)
        expert1 = self._get_expert1(commission_info)  # Только ФИО экспертов (обычных)
        expert2 = commission_info.get('experts_formatted', '')  # Эксперты через запятую с должностями

        print(f"memchair (ФИО председателя): '{memchair}'")
        print(f"members (все члены комиссии с ролями): '{members}'")
        print(f"expert1 (ФИО экспертов): '{expert1}'")
        print(f"expert2 (эксперты с должностями): '{expert2}'")

        # 5. Get settings data (provost, etc.)
        print("\n=== SETTINGS DATA ===")
        if settings_data:
            print(f"Settings data type: {type(settings_data)}, length: {len(settings_data)}")
            for i, field in enumerate(settings_data):
                if i < 10:  # Выводим первые 10 полей
                    field_type = type(field).__name__
                    field_str = str(field)[:100] if field else "None/Empty"
                    print(f"[{i}] {field_type}: {field_str}")
        else:
            print("No settings data provided")

        settings_info = self._extract_settings_data(settings_data) if settings_data else {}
        print(f"Provost (prorector): '{settings_info.get('provost', '')}'")
        print(f"Provost position: '{settings_info.get('provost_position', '')}'")
        print(f"POSTA (должность в именительном): '{settings_info.get('posta', '')}'")

        # 6. Get current date
        current_date = datetime.now().strftime('%d.%m.%Y')
        print(f"\n=== DATES ===")
        print(f"Current date: {current_date}")
        if application and len(application) > 10:
            app_date = application[10]
            print(f"Application date field [10]: {app_date}, Formatted: {self._format_date(app_date)}")

        # Replacements dictionary - для заключения
        replacements = {
            # Product/Program name
            '{{PRODUCT_NAME}}': product_name,
            '{PRODUCT_NAME}': product_name,
            'PRODUCTNAME': product_name,
            '{{PROGRAM_NAME}}': product_name,
            '{PROGRAM_NAME}': product_name,
            'PROGRAM_NAME': product_name,

            # Название заявки
            '{{APPLICATION_TITLE}}': application[1] if application and len(application) > 1 else '',
            '{APPLICATION_TITLE}': application[1] if application and len(application) > 1 else '',

            # Authors - в родительном падеже
            '{{AUTHORS_LIST}}': authors_data.get('authors_genitive_formatted', ''),
            '{AUTHORS_LIST}': authors_data.get('authors_genitive_formatted', ''),
            'AUTHORS': authors_data.get('authors_genitive_formatted', ''),

            # Authors - в именительном падеже
            '{{AUTHORS_NOMINATIVE}}': authors_data.get('authors_list_formatted', ''),
            '{AUTHORS_NOMINATIVE}': authors_data.get('authors_list_formatted', ''),

            # memchair - только ФИО председателя
            '{{MEMCHAIR}}': memchair,
            '{MEMCHAIR}': memchair,
            'MEMCHAIR': memchair,

            # members - все члены комиссии с ролями
            '{{MEMBERS}}': members,
            '{MEMBERS}': members,
            'MEMBERS': members,

            # Старые поля для совместимости
            '{{CHAIRMAN}}': commission_info.get('chairman', ''),
            '{CHAIRMAN}': commission_info.get('chairman', ''),
            'CHAIRMAN': commission_info.get('chairman', ''),

            # Все участники комиссии
            '{{COMMISSION_MEMBERS}}': members,
            '{COMMISSION_MEMBERS}': members,
            'COMMISSION_MEMBERS': members,

            # Для экспертов с форматированием
            '{{EXPERT_FORMATTED}}': commission_info.get('experts_formatted_multiline', ''),
            '{EXPERT_FORMATTED}': commission_info.get('experts_formatted_multiline', ''),
            'EXPERT_FORMATTED': commission_info.get('experts_formatted_multiline', ''),

            '{{EXPERT_1}}': expert1,
            '{EXPERT_1}': expert1,
            'EXPERT1': expert1,

            # expert2 - эксперты с должностями через запятую
            '{{EXPERT_2}}': expert2,
            '{EXPERT_2}': expert2,
            'EXPERT2': expert2,

            # expert3 - для совместимости
            '{{EXPERT_3}}': commission_info.get('expert_3_formatted', ''),
            '{EXPERT_3}': commission_info.get('expert_3_formatted', ''),
            'EXPERT_3': commission_info.get('expert_3_formatted', ''),

            # Все эксперты вместе
            '{{COMMISSION_EXPERTS}}': commission_info.get('all_experts_formatted', ''),
            '{COMMISSION_EXPERTS}': commission_info.get('all_experts_formatted', ''),
            'COMMISSION_EXPERTS': commission_info.get('all_experts_formatted', ''),

            # Rector
            '{{RECTOR}}': settings_info.get('rector', ''),
            '{RECTOR}': settings_info.get('rector', ''),
            'RECTOR': settings_info.get('rector', ''),

            # Provost - проректор (только фамилия и инициалы) - ИМЕНИТЕЛЬНЫЙ ПАДЕЖ
            '{{PROVOST}}': settings_info.get('provost', ''),
            '{PROVOST}': settings_info.get('provost', ''),
            'PROVOST': settings_info.get('provost', ''),

            # POSTA - должность проректора в именительном падеже
            '{{POSTA}}': settings_info.get('posta', ''),
            '{POSTA}': settings_info.get('posta', ''),
            'POSTA': settings_info.get('posta', ''),

            # Должность проректора (для совместимости)
            '{{PROVOST_POSITION}}': settings_info.get('provost_position', ''),
            '{PROVOST_POSITION}': settings_info.get('provost_position', ''),

            # Доверенность
            '{{POA_NUMBER}}': settings_info.get('poa_number', ''),
            '{POA_NUMBER}': settings_info.get('poa_number', ''),
            'POA_NUMBER': settings_info.get('poa_number', ''),

            '{{POA_DATE}}': settings_info.get('poa_date', ''),
            '{POA_DATE}': settings_info.get('poa_date', ''),
            'POA_DATE': settings_info.get('poa_date', ''),

            # Dates
            '{{CURRENT_DATE}}': current_date,
            '{CURRENT_DATE}': current_date,
            'CURRENT_DATE': current_date,

            # Application number and date
            '{{APP_NUMBER}}': application[0] if application and len(application) > 0 else '',
            '{APP_NUMBER}': application[0] if application and len(application) > 0 else '',
            'APP_NUMBER': application[0] if application and len(application) > 0 else '',

            '{{APP_DATE}}': self._format_date(application[10]) if application and len(
                application) > 10 else current_date,
            '{APP_DATE}': self._format_date(application[10]) if application and len(application) > 10 else current_date,
            'APP_DATE': self._format_date(application[10]) if application and len(application) > 10 else current_date,
        }

        print("\n=== REPLACEMENTS TO APPLY ===")
        print(f"MEMCHAIR (ФИО председателя): '{memchair}'")
        print(f"MEMBERS (все члены комиссии с ролями): '{members}'")
        print(f"EXPERT1 (ФИО экспертов): '{expert1}'")
        print(f"EXPERT2 (эксперты с должностями): '{expert2}'")
        print(f"PROVOST (проректор ФИО): '{settings_info.get('provost', '')}'")
        print(f"POSTA (должность проректора): '{settings_info.get('posta', '')}'")

        print("\nDEBUG: Applying replacements to document...")

        # Apply replacements
        self._replace_in_paragraphs(doc, replacements)
        self._replace_in_tables(doc, replacements)

        # Дополнительно обрабатываем специальные поля для экспертов
        self._process_expert_fields(doc, commission_info)

        print("\n" + "=" * 80)
        print("DEBUG: Conclusion template filled successfully")
        print("=" * 80)

    def _get_chairman1(self, commission_info):
        """Get chairman - only full name (ФИО только)"""
        if not commission_info.get('all_members'):
            return ''

        # Ищем председателя среди всех членов комиссии
        for member in commission_info['all_members']:
            if 'председатель' in member.get('role_name', '').lower():
                return member.get('full_name', '')
        return ''

    def _get_chairman2(self, commission_info):
        """Get chairman - full information (полная информация)"""
        if not commission_info.get('all_members'):
            return ''

        for member in commission_info['all_members']:
            if 'председатель' in member.get('role_name', '').lower():
                return member.get('formatted_line', '')
        return ''

    def _get_expert1(self, commission_info):
        """Get experts - only full names (только ФИО)"""
        if not commission_info.get('experts'):
            return ''

        expert_names = []
        for expert in commission_info['experts']:
            if expert.get('full_name'):
                expert_names.append(expert['full_name'])

        if not expert_names:
            return ''

        if len(expert_names) == 1:
            return expert_names[0]
        else:
            return f"{', '.join(expert_names[:-1])} и {expert_names[-1]}"

    def _get_expert2(self, commission_info):
        """Get all commission members with roles (все члены комиссии с ролями)"""
        if not commission_info.get('all_members'):
            return ''

        members_lines = []
        for member in commission_info['all_members']:
            role = member.get('role_name', '')
            full_name = member.get('full_name', '')
            position = member.get('position', '')

            if role and full_name:
                if position:
                    members_lines.append(f"{role} - {full_name}, {position}")
                else:
                    members_lines.append(f"{role} - {full_name}")

        return '\n'.join(members_lines)

    def _extract_product_name(self, application):
        """Extract product name from application"""
        if not application or not isinstance(application, tuple):
            print("WARNING: Invalid application data in _extract_product_name")
            return ''

        indices_to_try = [8, 7, 1]

        print(f"DEBUG _extract_product_name: Checking indices {indices_to_try}")

        for idx in indices_to_try:
            if len(application) > idx and application[idx]:
                name = str(application[idx]).strip()
                print(f"  Index {idx}: '{name}'")
                if name and name.lower() != 'admin' and not name.isdigit():
                    return name

        for idx in indices_to_try:
            if len(application) > idx and application[idx]:
                return str(application[idx]).strip()

        print("WARNING: No product name found in application")
        return ''

    def _extract_authors_data(self, application):
        """Extract authors data from application (для совместимости)"""
        return {
            'full_names_list': [],
            'authors_list_formatted': '',
            'authors_genitive_formatted': ''
        }

    def _extract_authors_from_list(self, authors_list):
        """Extract authors data from list of author tuples"""
        print(f"\nDEBUG _extract_authors_from_list: Processing {len(authors_list)} authors")

        result = {
            'full_names_list': [],
            'authors_list_formatted': '',
            'authors_genitive_formatted': ''
        }

        if not authors_list:
            print("  Empty authors list")
            return result

        for idx, author in enumerate(authors_list):
            print(f"\n  Author {idx}: type={type(author).__name__}")

            if isinstance(author, tuple):
                print(f"    Tuple length: {len(author)}")

                surname = ''
                name = ''
                patronymic = ''

                if len(author) >= 10:
                    surname = str(author[7]).strip() if author[7] else ''
                    name = str(author[8]).strip() if author[8] else ''
                    patronymic = str(author[9]).strip() if author[9] else ''

                    if surname and name:
                        print(f"    Extracted from [7,8,9]: {surname} {name} {patronymic}")

                if (not surname or not name) and len(author) >= 4:
                    test_surname = str(author[1]).strip() if author[1] else ''
                    test_name = str(author[2]).strip() if author[2] else ''
                    test_patronymic = str(author[3]).strip() if len(author) > 3 and author[3] else ''

                    if test_surname and test_name and any(c.isalpha() for c in test_surname):
                        surname, name, patronymic = test_surname, test_name, test_patronymic
                        print(f"    Extracted from [1,2,3]: {surname} {name} {patronymic}")

                if surname and name:
                    full_name = f"{surname} {name} {patronymic}".strip()
                    result['full_names_list'].append(full_name)
                    print(f"    ✓ Added author: {full_name}")
                else:
                    print(f"    ✗ Could not extract author name")

            elif isinstance(author, str):
                full_name = author.strip()
                if full_name:
                    result['full_names_list'].append(full_name)
                    print(f"    ✓ String author: {full_name}")
            else:
                print(f"    ✗ Unsupported author type: {type(author)}")

        if result['full_names_list']:
            print(f"\n  Processing {len(result['full_names_list'])} authors:")

            if len(result['full_names_list']) == 1:
                result['authors_list_formatted'] = result['full_names_list'][0]
            else:
                authors_str = ', '.join(result['full_names_list'][:-1])
                result['authors_list_formatted'] = f"{authors_str} и {result['full_names_list'][-1]}"

            genitive_list = []
            for author_name in result['full_names_list']:
                parts = author_name.strip().split()
                if len(parts) >= 2:
                    surname = parts[0]
                    name = parts[1]
                    patronymic = parts[2] if len(parts) > 2 else ''

                    genitive_surname = self._to_genitive_surname(surname)

                    initials = ''
                    if name:
                        initials += name[0].upper() + '.'
                    if patronymic:
                        initials += patronymic[0].upper() + '.'

                    genitive_name = f"{genitive_surname} {initials}".strip()
                    genitive_list.append(genitive_name)
                    print(f"    '{author_name}' -> '{genitive_name}'")
                else:
                    genitive_list.append(author_name)

            if len(genitive_list) == 1:
                result['authors_genitive_formatted'] = genitive_list[0]
            else:
                genitive_str = ', '.join(genitive_list[:-1])
                result['authors_genitive_formatted'] = f"{genitive_str} и {genitive_list[-1]}"
        else:
            print("\n  No authors extracted")

        print(f"\n  Final: {len(result['full_names_list'])} authors")
        print(f"  Genitive: '{result['authors_genitive_formatted']}'")

        return result

    def _to_genitive_surname(self, surname):
        """Convert surname to genitive case (simplified)"""
        if not surname:
            return ''

        surname_lower = surname.strip().lower()

        if surname_lower.endswith('ова') or surname_lower.endswith('ева') or surname_lower.endswith('ина'):
            result = surname[:-1] + 'ой'
        elif surname_lower.endswith('ая'):
            result = surname[:-2] + 'ой'
        elif surname_lower.endswith('ская'):
            result = surname[:-4] + 'ской'
        elif surname_lower.endswith('цкая'):
            result = surname[:-4] + 'цкой'
        elif surname_lower.endswith('кая'):
            result = surname[:-3] + 'кой'
        elif surname_lower.endswith('га') or surname_lower.endswith('ка'):
            result = surname[:-1] + 'и'
        elif surname_lower.endswith('я'):
            result = surname[:-1] + 'и'
        elif surname_lower.endswith('ов') or surname_lower.endswith('ев') or surname_lower.endswith('ин'):
            result = surname + 'а'
        elif surname_lower.endswith('ий') or surname_lower.endswith('ый'):
            result = surname[:-2] + 'ого'
        elif surname_lower.endswith('ой'):
            result = surname[:-2] + 'ого'
        else:
            result = surname + 'а'

        return result

    def _extract_commission_data(self, commission_data):
        """Extract commission experts data"""
        print(f"\nDEBUG _extract_commission_data: Processing commission data")
        print(f"  Type: {type(commission_data)}")
        print(f"  Length: {len(commission_data) if commission_data else 0}")

        if commission_data and len(commission_data) > 0:
            print(f"  First item: {commission_data[0]}")
            print(f"  First item type: {type(commission_data[0])}")
            if isinstance(commission_data[0], tuple):
                print(f"  First tuple length: {len(commission_data[0])}")
                for i, field in enumerate(commission_data[0]):
                    print(f"    [{i}] {type(field).__name__}: {field}")

        result = {
            'chairman': '',
            'experts': [],
            'all_members': [],
            'expert_1': '',
            'expert_2': '',
            'expert_3': '',
            'experts_formatted': '',
            'experts_formatted_multiline': '',
            'expert_1_formatted': '',
            'expert_2_formatted': '',
            'expert_2_detailed': '',
            'expert_3_formatted': '',
            'all_experts_formatted': '',
            'all_commission_members': ''
        }

        if not commission_data:
            print("  No commission data provided")
            return result

        experts_list = commission_data if isinstance(commission_data, list) else []

        chairman_list = []
        regular_experts = []
        all_members = []

        for idx, expert in enumerate(experts_list):
            print(f"\n  Processing expert {idx}")

            if isinstance(expert, tuple):
                print(f"    Tuple length: {len(expert)}")

                if len(expert) >= 10:
                    surname_idx, name_idx, patronymic_idx, position_idx, role_idx = None, None, None, None, None

                    if isinstance(expert[5], str) and isinstance(expert[6], str):
                        surname_idx, name_idx, patronymic_idx, position_idx, role_idx = 5, 6, 7, 8, 9
                        print(f"    Using new structure [5,6,7,8,9]")

                    elif len(expert) > 5 and isinstance(expert[1], str) and isinstance(expert[2], str):
                        surname_idx, name_idx, patronymic_idx, position_idx, role_idx = 1, 2, 3, 4, 5
                        print(f"    Using old structure [1,2,3,4,5]")

                    if surname_idx is not None:
                        surname = str(expert[surname_idx]).strip() if expert[surname_idx] else ''
                        name = str(expert[name_idx]).strip() if expert[name_idx] else ''
                        patronymic = str(expert[patronymic_idx]).strip() if patronymic_idx < len(expert) and expert[
                            patronymic_idx] else ''
                        position = str(expert[position_idx]).strip() if position_idx < len(expert) and expert[
                            position_idx] else ''
                        role_name = str(expert[role_idx]).lower() if role_idx < len(expert) and expert[role_idx] else ''

                        print(f"    Extracted: {surname} {name} {patronymic}, position: {position}, role: {role_name}")

                        full_name = f"{surname} {name} {patronymic}".strip()

                        initials = ''
                        if name:
                            initials += name[0].upper() + '.'
                        if patronymic:
                            initials += patronymic[0].upper() + '.'

                        signature = f"{initials} {surname}".strip() if surname else ''

                        expert_info = {
                            'full_name': full_name,
                            'surname': surname,
                            'name': name,
                            'patronymic': patronymic,
                            'position': position,
                            'role_name': role_name,
                            'initials': initials,
                            'signature': signature,
                            'formatted_line': f"{full_name}, {position}" if position else full_name,
                            'with_role': f"{role_name} - {full_name}, {position}" if position else f"{role_name} - {full_name}"
                        }

                        if 'председатель' in role_name:
                            chairman_list.append(expert_info)
                            all_members.append(expert_info)
                            print(f"    Added as chairman: {full_name}")
                        elif 'эксперт' in role_name or 'секретарь' in role_name:
                            regular_experts.append(expert_info)
                            all_members.append(expert_info)
                            print(f"    Added as regular expert: {full_name}")
                        else:
                            all_members.append(expert_info)
                            print(f"    Added with role '{role_name}': {full_name}")
                    else:
                        print(f"    Could not determine data structure")
                else:
                    print(f"    Tuple too short ({len(expert)} < 10), skipping")
            else:
                print(f"    Not a tuple, type: {type(expert)}")

        if chairman_list:
            chairman = chairman_list[0]
            result['chairman'] = chairman['formatted_line']
            print(f"\n  Chairman found: {result['chairman']}")
        else:
            print("\n  No chairman found")

        if regular_experts:
            result['experts'] = regular_experts
            print(f"  Found {len(regular_experts)} regular experts")

            formatted_list = [exp['formatted_line'] for exp in regular_experts]
            if len(formatted_list) == 1:
                result['experts_formatted'] = formatted_list[0]
            elif len(formatted_list) > 1:
                experts_str = ', '.join(formatted_list[:-1])
                result['experts_formatted'] = f"{experts_str} и {formatted_list[-1]}"

            multiline_parts = []
            for i, expert in enumerate(regular_experts):
                multiline_parts.append(expert['formatted_line'])
                multiline_parts.append("________________")
                multiline_parts.append(expert['signature'] if expert['signature'] else "(подпись)")
                if i < len(regular_experts) - 1:
                    multiline_parts.append("")

            result['experts_formatted_multiline'] = '\n'.join(multiline_parts)
            result['expert_1_formatted'] = result['experts_formatted_multiline']

        if all_members:
            result['all_members'] = all_members

            all_members_parts = []
            for i, member in enumerate(all_members):
                all_members_parts.append(member['with_role'])

            result['expert_2_detailed'] = '\n'.join(all_members_parts)
            result['all_commission_members'] = result['expert_2_detailed']

            all_multiline = []
            for i, member in enumerate(all_members):
                all_multiline.append(member['formatted_line'])
                all_multiline.append("________________")
                all_multiline.append(member['signature'] if member['signature'] else "(подпись)")
                if i < len(all_members) - 1:
                    all_multiline.append("")

            result['expert_2_formatted'] = '\n'.join(all_multiline)

        all_experts = []
        if result['chairman']:
            all_experts.append(f"председатель - {result['chairman']}")
        if result['experts_formatted']:
            all_experts.append(f"эксперты - {result['experts_formatted']}")

        result['all_experts_formatted'] = ', '.join(all_experts)

        print(f"\n  Final commission info:")
        print(f"    Chairman: {result['chairman']}")
        print(f"    Regular experts: {len(regular_experts)}")
        print(f"    All members: {len(all_members)}")

        return result

    def _extract_settings_data(self, settings_data):
        """Extract settings data - проректор и должность в ИМЕНИТЕЛЬНОМ падеже"""
        print(f"\nDEBUG _extract_settings_data: Processing settings")

        result = {
            'rector': '',
            'rector_full': '',
            'provost': '',  # Проректор (только фамилия и инициалы) - именительный падеж
            'provost_position': '',  # Должность проректора - именительный падеж
            'posta': '',  # Должность проректора для плейсхолдера POSTA - именительный падеж
            'assignee': '',
            'poa_number': '',
            'poa_date': '',
            'sum_value': ''
        }

        if not settings_data:
            print("  No settings data")
            return result

        print(f"  Settings data type: {type(settings_data)}, length: {len(settings_data)}")

        if isinstance(settings_data, tuple):
            # Полное ФИО ректора (индекс 5)
            if len(settings_data) > 5 and settings_data[5]:
                full_rector_name = str(settings_data[5]).strip()
                result['rector_full'] = full_rector_name
                print(f"    Found full rector name at index 5: '{full_rector_name}'")

                result['rector'] = self._convert_to_initials(full_rector_name)
                print(f"    Converted to initials: '{result['rector']}'")

            # Должность проректора (индекс 8) - ИМЕНИТЕЛЬНЫЙ ПАДЕЖ
            if len(settings_data) > 8 and settings_data[8]:
                provost_position = str(settings_data[8]).strip()
                result['provost_position'] = provost_position
                result['posta'] = provost_position  # Для плейсхолдера POSTA
                print(f"    Found provost position at index 8: '{provost_position}'")
                print(f"    POSTA (nominative): '{result['posta']}'")

            # Имя проректора (индекс 7) - ИМЕНИТЕЛЬНЫЙ ПАДЕЖ
            if len(settings_data) > 7 and settings_data[7]:
                provost_name = str(settings_data[7]).strip()
                # Преобразуем в фамилию и инициалы
                result['provost'] = self._convert_to_initials(provost_name)
                print(f"    Found provost name at index 7: '{provost_name}'")
                print(f"    Converted to initials: '{result['provost']}'")

            # Номер доверенности (индекс 6)
            if len(settings_data) > 6 and settings_data[6]:
                result['poa_number'] = str(settings_data[6]).strip()
                print(f"    Found POA number at index 6: '{result['poa_number']}'")

            # Дата доверенности (индекс 7) - если есть отдельное поле
            if len(settings_data) > 7 and settings_data[7]:
                # Проверяем, что это не строка с именем
                if not isinstance(settings_data[7], str) or len(str(settings_data[7])) > 20:
                    result['poa_date'] = self._format_date(settings_data[7])
                    print(f"    Found POA date: '{settings_data[7]}' -> '{result['poa_date']}'")

            # Сумма (индекс 9)
            if len(settings_data) > 9 and settings_data[9]:
                result['sum_value'] = str(settings_data[9])
                print(f"    Found sum at index 9: '{result['sum_value']}'")

        print(f"\n  Extracted settings:")
        print(f"    Provost (initials): '{result['provost']}'")
        print(f"    Provost position: '{result['provost_position']}'")
        print(f"    POSTA: '{result['posta']}'")
        print(f"    Power of attorney: {result['poa_number']} от {result['poa_date']}")
        print(f"    Sum: {result['sum_value']}")

        return result

    def _convert_to_initials(self, full_name):
        """Convert full name to surname and initials (Фамилия И.О.)"""
        if not full_name:
            return ''

        full_name = ' '.join(full_name.strip().split())
        parts = full_name.split()

        if len(parts) >= 3:
            surname = parts[0]
            name = parts[1]
            patronymic = parts[2]

            name_init = name[0].upper() + '.' if name else ''
            patronymic_init = patronymic[0].upper() + '.' if patronymic else ''

            return f"{surname} {name_init}{patronymic_init}"
        elif len(parts) == 2:
            surname = parts[0]
            name = parts[1]
            name_init = name[0].upper() + '.' if name else ''
            return f"{surname} {name_init}"
        elif len(parts) == 1:
            return parts[0]
        else:
            return full_name

    def _set_single_line_spacing(self, paragraph):
        """Установка одинарного междустрочного интервала"""
        paragraph_format = paragraph.paragraph_format
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        paragraph_format.line_spacing = 1.0
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)

    def _process_expert_fields(self, doc, commission_info):
        """Специальная обработка полей для экспертов и председателя"""
        experts = commission_info.get('experts', [])

        chairman = None
        if commission_info.get('all_members'):
            for member in commission_info['all_members']:
                if 'председатель' in member.get('role_name', '').lower():
                    chairman = member
                    break

        self._process_memchair_field(doc, chairman)

        if experts:
            self._process_expert_field(doc, experts)

        self._process_commission_members_field(doc, commission_info)

    def _process_memchair_field(self, doc, chairman):
        """Обработка поля MEMCHAIR (председатель комиссии)"""
        if not chairman:
            return

        full_name = chairman.get('full_name', '')

        if not full_name:
            return

        name_with_spaces = f"      {full_name}"

        for paragraph in doc.paragraphs:
            if '{{MEMCHAIR}}' in paragraph.text or '{MEMCHAIR}' in paragraph.text or 'MEMCHAIR' in paragraph.text:
                paragraph.clear()
                self._set_single_line_spacing(paragraph)

                if full_name:
                    name_run = paragraph.add_run(f"{name_with_spaces}")
                    name_run.underline = True
                    name_run.font.name = 'Times New Roman'
                    name_run.font.size = Pt(11)

                paragraph.add_run("\n")

                indent = " " * 50
                signature_run = paragraph.add_run(f"{indent}(подпись, инициалы и фамилия)")
                signature_run.font.name = 'Times New Roman'
                signature_run.font.size = Pt(9)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{MEMCHAIR}}' in paragraph.text or '{MEMCHAIR}' in paragraph.text or 'MEMCHAIR' in paragraph.text:
                            paragraph.clear()
                            self._set_single_line_spacing(paragraph)

                            if full_name:
                                name_run = paragraph.add_run(f"{name_with_spaces}")
                                name_run.underline = True
                                name_run.font.name = 'Times New Roman'
                                name_run.font.size = Pt(11)

                            paragraph.add_run("\n")

                            indent = " " * 50
                            signature_run = paragraph.add_run(f"{indent}(подпись, инициалы и фамилия)")
                            signature_run.font.name = 'Times New Roman'
                            signature_run.font.size = Pt(9)

    def _process_expert_field(self, doc, experts):
        """Обработка поля EXPERT (обычные эксперты)"""
        for paragraph in doc.paragraphs:
            if '{{EXPERT}}' in paragraph.text or '{EXPERT}' in paragraph.text or 'EXPERT' in paragraph.text:
                paragraph.clear()
                self._set_single_line_spacing(paragraph)

                for i, expert in enumerate(experts):
                    full_name = expert.get('full_name', '')

                    if not full_name:
                        continue

                    name_with_spaces = f"                               {full_name}"

                    name_run = paragraph.add_run(f"{name_with_spaces}")
                    name_run.underline = True
                    name_run.font.name = 'Times New Roman'
                    name_run.font.size = Pt(11)

                    paragraph.add_run("\n")

                    signature_run = paragraph.add_run(f"     (подпись, инициалы и фамилия)")
                    signature_run.font.name = 'Times New Roman'
                    signature_run.font.size = Pt(9)

                    if i < len(experts) - 1:
                        paragraph.add_run("\n")
                        paragraph.add_run("\n")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{EXPERT}}' in paragraph.text or '{EXPERT}' in paragraph.text or 'EXPERT' in paragraph.text:
                            paragraph.clear()
                            self._set_single_line_spacing(paragraph)

                            for i, expert in enumerate(experts):
                                full_name = expert.get('full_name', '')

                                if not full_name:
                                    continue

                                name_with_spaces = f"      {full_name}"

                                name_run = paragraph.add_run(f"{name_with_spaces}")
                                name_run.underline = True
                                name_run.font.name = 'Times New Roman'
                                name_run.font.size = Pt(11)

                                paragraph.add_run("\n")

                                indent = " " * 50
                                signature_run = paragraph.add_run(f"{indent}(подпись, инициалы и фамилия)")
                                signature_run.font.name = 'Times New Roman'
                                signature_run.font.size = Pt(9)

                                if i < len(experts) - 1:
                                    paragraph.add_run("\n")

    def _process_commission_members_field(self, doc, commission_info):
        """Обработка поля COMMISSION_MEMBERS (все члены комиссии)"""
        all_members = commission_info.get('all_members', [])

        if not all_members:
            return

        for paragraph in doc.paragraphs:
            if '{{COMMISSION_MEMBERS}}' in paragraph.text or 'COMMISSION' in paragraph.text or 'COMMISSION_MEMBERS' in paragraph.text:
                paragraph.clear()
                self._set_single_line_spacing(paragraph)

                for i, member in enumerate(all_members):
                    full_name = member.get('full_name', '')
                    role_name = member.get('role_name', '')

                    if not full_name:
                        continue

                    name_with_spaces = f"      {full_name}"

                    if role_name and role_name.lower() != 'эксперт':
                        role_run = paragraph.add_run(f"{role_name} - ")
                        role_run.font.name = 'Times New Roman'
                        role_run.font.size = Pt(11)

                    name_run = paragraph.add_run(f"{name_with_spaces}")
                    name_run.underline = True
                    name_run.font.name = 'Times New Roman'
                    name_run.font.size = Pt(11)

                    paragraph.add_run("\n")

                    indent = " " * 50
                    signature_run = paragraph.add_run(f"{indent}(подпись, инициалы и фамилия)")
                    signature_run.font.name = 'Times New Roman'
                    signature_run.font.size = Pt(9)

                    if i < len(all_members) - 1:
                        paragraph.add_run("\n")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{COMMISSION_MEMBERS}}' in paragraph.text or 'COMMISSION' in paragraph.text or 'COMMISSION_MEMBERS' in paragraph.text:
                            paragraph.clear()
                            self._set_single_line_spacing(paragraph)

                            for i, member in enumerate(all_members):
                                full_name = member.get('full_name', '')
                                role_name = member.get('role_name', '')

                                if not full_name:
                                    continue

                                name_with_spaces = f"      {full_name}"

                                if role_name and role_name.lower() != 'эксперт':
                                    role_run = paragraph.add_run(f"{role_name} - ")
                                    role_run.font.name = 'Times New Roman'
                                    role_run.font.size = Pt(11)

                                name_run = paragraph.add_run(f"{name_with_spaces}")
                                name_run.underline = True
                                name_run.font.name = 'Times New Roman'
                                name_run.font.size = Pt(11)

                                paragraph.add_run("\n")

                                indent = " " * 50
                                signature_run = paragraph.add_run(f"{indent}(подпись, инициалы и фамилия)")
                                signature_run.font.name = 'Times New Roman'
                                signature_run.font.size = Pt(9)

                                if i < len(all_members) - 1:
                                    paragraph.add_run("\n")

    def _format_date(self, date_value):
        """Format date safely"""
        try:
            if isinstance(date_value, (datetime, date)):
                return date_value.strftime('%d.%m.%Y')
            elif isinstance(date_value, str):
                for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
                    try:
                        dt = datetime.strptime(date_value, fmt)
                        return dt.strftime('%d.%m.%Y')
                    except:
                        continue
            return str(date_value)
        except:
            return str(date_value) if date_value else ''

    def _replace_in_paragraphs(self, doc, replacements):
        """Replace text in paragraphs"""
        for paragraph in doc.paragraphs:
            if paragraph.text:
                for old, new in replacements.items():
                    if old in paragraph.text:
                        for run in paragraph.runs:
                            if old in run.text:
                                run.text = run.text.replace(old, new)

    def _replace_in_tables(self, doc, replacements):
        """Replace text in tables"""
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text:
                            for old, new in replacements.items():
                                if old in paragraph.text:
                                    for run in paragraph.runs:
                                        if old in run.text:
                                            run.text = run.text.replace(old, new)