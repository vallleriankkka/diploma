import os
import io
from datetime import datetime, date
from docx import Document
from flask import current_app, Response
from flask import session, flash, redirect, url_for


class SoglGenerator:
    """Document generator for consent form"""

    def __init__(self, template_name='sogl_na_obr_pers_dan.docx'):
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

    def generate_for_all(self, application, application_authors, author_dao=None):
        """Generate single document for all authors"""
        try:
            if not os.path.exists(self.template_path):
                print(f"ERROR: Template not found: {self.template_path}")
                return None

            doc = Document(self.template_path)

            # Получаем полные данные авторов
            full_authors_data = []
            if application_authors and author_dao:
                for app_author in application_authors:
                    # Извлекаем ID автора из application_author
                    # В application_author author_id обычно на индексе 2 или 3
                    author_id = None
                    if len(app_author) > 2 and app_author[2]:  # пробуем индекс 2
                        author_id = app_author[2]
                    elif len(app_author) > 3 and app_author[3]:  # пробуем индекс 3
                        author_id = app_author[3]

                    if author_id:
                        print(f"DEBUG: Fetching author data for ID: {author_id}")
                        author_data = author_dao.get_by_id(author_id)
                        if author_data:
                            full_authors_data.append(author_data)
                        else:
                            print(f"DEBUG: No author data found for ID: {author_id}")
                    else:
                        print(f"DEBUG: No author_id found in app_author")

            print(f"DEBUG: Got {len(full_authors_data)} full author records")
            self._fill_template(doc, application, full_authors_data)
            return doc

        except Exception as e:
            print(f"ERROR in generate: {e}")
            import traceback
            traceback.print_exc()
            return None

    def generate_for_each_author(self, application, application_authors, author_dao=None):
        """Generate separate document for each author"""
        if not application_authors:
            return []

        documents = []

        # Получаем полные данные авторов
        full_authors_data = []
        if author_dao:
            for app_author in application_authors:
                # Извлекаем ID автора из application_author
                # В application_author author_id обычно на индексе 2 или 3
                author_id = None
                if len(app_author) > 2 and app_author[2]:  # пробуем индекс 2
                    author_id = app_author[2]
                elif len(app_author) > 3 and app_author[3]:  # пробуем индекс 3
                    author_id = app_author[3]

                if author_id:
                    print(f"DEBUG: Fetching author data for ID: {author_id}")
                    author_data = author_dao.get_by_id(author_id)
                    if author_data:
                        full_authors_data.append(author_data)
                    else:
                        print(f"DEBUG: No author data found for ID: {author_id}")
                else:
                    print(f"DEBUG: No author_id found in app_author")

        print(f"DEBUG: Got {len(full_authors_data)} full author records")

        for author in full_authors_data:
            try:
                if not os.path.exists(self.template_path):
                    print(f"ERROR: Template not found: {self.template_path}")
                    continue

                # Create a new document for each author
                doc = Document(self.template_path)

                # Fill template with single author
                self._fill_template_for_single_author(doc, application, author)

                documents.append(doc)

            except Exception as e:
                print(f"ERROR generating document for author: {e}")
                import traceback
                traceback.print_exc()
                continue

        return documents

    def _fill_template(self, doc, application, authors):
        """Fill template with data for all authors"""
        print("DEBUG: Filling template for all authors...")

        # Get product name
        product_name = self._extract_product_name(application)
        print(f"DEBUG: Product name: {product_name}")

        # Get authors data
        authors_data = self._extract_authors_data(authors)

        print(f"DEBUG: Found {len(authors) if authors else 0} authors")

        # Replacements dictionary
        replacements = {
            '{{PRODUCT_NAME}}': product_name,
            '{PRODUCT_NAME}': product_name,
            'PRODUCTNAME': product_name,

            '{{PROGRAM_NAME}}': product_name,
            '{PROGRAM_NAME}': product_name,
            'PROGRAM_NAME': product_name,

            '{{AUTHOR_FULL_NAME}}': authors_data.get('full_names', ''),
            '{AUTHOR_FULL_NAME}': authors_data.get('full_names', ''),
            'FULLNAME': authors_data.get('full_names', ''),

            '{{AUTHORS_NAMES}}': authors_data.get('full_names', ''),
            '{AUTHORS_NAMES}': authors_data.get('full_names', ''),
            'AUTHORS_NAMES': authors_data.get('full_names', ''),

            '{{ADDRESS_FULL}}': authors_data.get('addresses_full', ''),
            '{ADDRESS_FULL}': authors_data.get('addresses_full', ''),
            'ADDRESS_FULL': authors_data.get('addresses_full', ''),

            '{{ADDRESS}}': authors_data.get('addresses_full', ''),
            '{ADDRESS}': authors_data.get('addresses_full', ''),
            'ADDRESS': authors_data.get('addresses_full', ''),

            '{{PASSPORT_FULL}}': authors_data.get('passport_full', ''),
            '{PASSPORT_FULL}': authors_data.get('passport_full', ''),
            'PASSPORT_FULL': authors_data.get('passport_full', ''),

            '{{PASSPORT_DATA}}': authors_data.get('passport_full', ''),
            '{PASSPORT_DATA}': authors_data.get('passport_full', ''),
            'PASSPORT': authors_data.get('passport_full', ''),
        }

        print("DEBUG: Applying replacements...")

        # Apply replacements
        self._replace_in_paragraphs(doc, replacements)
        self._replace_in_tables(doc, replacements)

        print("DEBUG: Template filled successfully")

    def _fill_template_for_single_author(self, doc, application, author):
        """Fill template with data for single author"""
        print("DEBUG: Filling template for single author...")

        # Get product name
        product_name = self._extract_product_name(application)
        print(f"DEBUG: Product name: {product_name}")

        # Get single author data
        author_data = self._extract_single_author_data(author)

        print("DEBUG: Single author data extracted")

        # Replacements dictionary for single author
        replacements = {
            '{{PRODUCT_NAME}}': product_name,
            '{PRODUCT_NAME}': product_name,
            'PRODUCTNAME': product_name,

            '{{PROGRAM_NAME}}': product_name,
            '{PROGRAM_NAME}': product_name,
            'PROGRAM_NAME': product_name,

            '{{AUTHOR_FULL_NAME}}': author_data.get('full_name', ''),
            '{AUTHOR_FULL_NAME}': author_data.get('full_name', ''),
            'FULLNAME': author_data.get('full_name', ''),

            '{{AUTHORS_NAMES}}': author_data.get('full_name', ''),
            '{AUTHORS_NAMES}': author_data.get('full_name', ''),
            'AUTHORS_NAMES': author_data.get('full_name', ''),

            '{{ADDRESS_FULL}}': author_data.get('address', ''),
            '{ADDRESS_FULL}': author_data.get('address', ''),
            'ADDRESS_FULL': author_data.get('address', ''),

            '{{ADDRESS}}': author_data.get('address', ''),
            '{ADDRESS}': author_data.get('address', ''),
            'ADDRESS': author_data.get('address', ''),

            '{{PASSPORT_FULL}}': author_data.get('passport', ''),
            '{PASSPORT_FULL}': author_data.get('passport', ''),
            'PASSPORT_FULL': author_data.get('passport', ''),

            '{{PASSPORT_DATA}}': author_data.get('passport', ''),
            '{PASSPORT_DATA}': author_data.get('passport', ''),
            'PASSPORT': author_data.get('passport', ''),
        }

        print("DEBUG: Applying replacements for single author...")

        # Apply replacements
        self._replace_in_paragraphs(doc, replacements)
        self._replace_in_tables(doc, replacements)

        print("DEBUG: Single author template filled successfully")

    def _extract_product_name(self, application):
        """Extract product name"""
        if not application or not isinstance(application, tuple):
            return ''

        # Try index 7 (product name) as fallback
        if len(application) > 8 and application[8]:
            name = str(application[8]).strip()
            if name:
                return name

        return ''

    def _extract_authors_data(self, authors):
        """Extract data for all authors"""
        result = {
            'full_names': '',
            'addresses_full': '',
            'passport_full': ''
        }

        if not authors:
            return result

        full_names = []
        addresses = []
        passport_list = []

        for i, author in enumerate(authors):
            if not isinstance(author, tuple):
                continue

            # Extract data for single author
            author_data = self._extract_single_author_data(author)

            if author_data.get('full_name'):
                full_names.append(author_data['full_name'])

            if author_data.get('address'):
                addresses.append(author_data['address'])

            if author_data.get('passport'):
                passport_list.append(author_data['passport'])

        # Combine results
        result['full_names'] = ', '.join(full_names) if full_names else ''
        result['addresses_full'] = ', '.join(addresses) if addresses else ''
        result['passport_full'] = '; '.join(passport_list) if passport_list else ''

        return result

    def _extract_single_author_data(self, author):
        """Extract data for single author from author table"""
        if not isinstance(author, tuple):
            return {}

        result = {
            'full_name': '',
            'address': '',
            'passport': ''
        }

        print(f"\n{'=' * 50}")
        print("DEBUG: Processing single author (FULL author data)")
        print(f"DEBUG: Author tuple length: {len(author)}")

        # Выводим всю структуру для отладки
        for i, val in enumerate(author):
            field_type = type(val).__name__ if val is not None else 'None'
            if isinstance(val, str):
                print(f"  [{i:2d}] = '{val}' ({field_type})")
            else:
                print(f"  [{i:2d}] = {val} ({field_type})")

        print('=' * 50)

        # 1. Extract full name - в таблице author индексы другие!
        # Судя по структуре таблицы author:
        # 0 - id, 1 - surname, 2 - name, 3 - patronymic, 4 - birth_date, 5 - address
        surname = ''
        name = ''
        patronymic = ''

        if len(author) > 1 and author[1]:  # surname
            surname = str(author[1]).strip()
        if len(author) > 2 and author[2]:  # name
            name = str(author[2]).strip()
        if len(author) > 3 and author[3]:  # patronymic
            patronymic = str(author[3]).strip()

        result['full_name'] = f"{surname} {name} {patronymic}".strip()
        print(f"DEBUG: Full name extracted: '{result['full_name']}'")

        # 2. Extract address - индекс 5 в таблице author
        address = ''
        if len(author) > 5 and author[5]:
            address = str(author[5]).strip()
            print(f"DEBUG: Address extracted: '{address[:50]}...'")

        result['address'] = address

        # 3. Extract passport - новые индексы для таблицы author
        result['passport'] = self._extract_passport_from_author(author)
        print(f"DEBUG: Passport extracted: '{result['passport']}'")

        return result

    def _extract_passport_from_author(self, author):
        """Extract passport data from author table"""
        if not isinstance(author, tuple):
            return ''

        passport_parts = []

        # Судя по структуре таблицы author:
        # 7 - passport (серия и номер)
        # 8 - passport_issue_date
        # 9 - passport_issued_by
        # 10 - passport_issued_code

        if len(author) > 6 and author[6]:  # passport
            passport = str(author[6]).strip()
            if passport:
                passport_parts.append(f"паспорт {passport}")

        issued_by = ''
        if len(author) > 16 and author[16]:  # passport_issued_by
            issued_by = str(author[16]).strip()

        issue_date = ''
        if len(author) > 15 and author[15]:  # passport_issue_date
            issue_date = self._format_date(author[15])

        issued_code = ''
        if len(author) > 17 and author[17]:  # passport_issued_code
            issued_code = str(author[17]).strip()

        # Format passport string
        result = ''
        if passport_parts:
            result = ', '.join(passport_parts)

        if issued_by:
            if result:
                result += f", выдан: {issued_by}"
            else:
                result = f"выдан: {issued_by}"

        if issue_date:
            if result:
                result += f", {issue_date}"
            else:
                result = issue_date

        if issued_code:
            if result:
                result += f" (код подразделения: {issued_code})"
            else:
                result = f"код подразделения: {issued_code}"

        return result if result else 'не указаны'

    def _extract_passport(self, author):
        """Old method - kept for compatibility"""
        return self._extract_passport_from_author(author)

    def _format_date(self, date_value):
        """Format date safely"""
        try:
            if isinstance(date_value, (datetime, date)):
                return date_value.strftime('%d.%m.%Y')
            elif isinstance(date_value, str):
                # Try different date formats
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
