import os
from datetime import datetime, date
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from flask import current_app


class DistributionAgreementGenerator:
    """Document generator for distribution agreement between authors"""

    def __init__(self, template_name='distributionagreement_generator.docx'):
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

    def generate(self, application, authors_data=None):
        """Generate distribution agreement document"""
        try:
            if not os.path.exists(self.template_path):
                print(f"ERROR: Template not found: {self.template_path}")
                return None

            doc = Document(self.template_path)
            self._fill_template(doc, application, authors_data)
            return doc

        except Exception as e:
            print(f"ERROR in generate: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _fill_template(self, doc, application, authors_data):
        """Fill template with data"""
        print("\n" + "=" * 80)
        print("DEBUG: START Filling distribution agreement template...")
        print("=" * 80)

        # 1. Get product name
        product_name = self._extract_product_name(application)
        print(f"\n=== PRODUCT NAME ===")
        print(f"Extracted: '{product_name}'")

        # 2. Get authors data with position/status information
        print("\n=== AUTHORS DATA ===")
        if authors_data is None:
            print("No authors_data passed, extracting from application")
            authors_data = self._extract_authors_data(application)
        else:
            print(
                f"Authors data passed directly, type: {type(authors_data)}, length: {len(authors_data) if hasattr(authors_data, '__len__') else 'N/A'}")
            authors_data = self._extract_authors_from_list(authors_data)

        print(f"Found {len(authors_data['full_info_list']) if authors_data else 0} authors")
        for i, author_info in enumerate(authors_data.get('full_info_list', [])):
            print(f"Author {i + 1}: {author_info.get('full_name')} - {author_info.get('status_for_table', 'N/A')}")

        # 3. Get NIR info from application - ИСПРАВЛЕНО
        nir_info = self._extract_nir_info(application)
        print(f"\n=== NIR INFO ===")
        print(f"Extracted: '{nir_info}'")

        # 4. Get current date
        current_date = datetime.now().strftime('%d.%m.%Y')
        print(f"\n=== DATES ===")
        print(f"Current date: {current_date}")

        # 5. Get application info
        app_number = application[0] if application and len(application) > 0 else ''
        app_date = self._format_date(application[10]) if application and len(application) > 10 else current_date

        # 6. Create replacements dictionary
        replacements = {
            # Product/Program name
            '{{PRODUCT_NAME}}': product_name,
            '{PRODUCT_NAME}': product_name,
            'PRODUCTNAME': product_name,
            '{{PROGRAM_NAME}}': product_name,
            '{PROGRAM_NAME}': product_name,
            'PROGRAM_NAME': product_name,

            # NIR information (extracted from application)
            '{{NIR}}': nir_info,
            '{NIR}': nir_info,
            'NIR': nir_info,
            '{{ORGANIZATION}}': nir_info,
            '{ORGANIZATION}': nir_info,
            'ORGANIZATION': nir_info,
            '{{DEPARTMENT}}': nir_info,
            '{DEPARTMENT}': nir_info,
            'DEPARTMENT': nir_info,

            # Current date
            '{{CURRENT_DATE}}': current_date,
            '{CURRENT_DATE}': current_date,
            'CURRENT_DATE': current_date,

            # Application number and date
            '{{APP_NUMBER}}': app_number,
            '{APP_NUMBER}': app_number,
            'APP_NUMBER': app_number,

            '{{APP_DATE}}': app_date,
            '{APP_DATE}': app_date,
            'APP_DATE': app_date,
        }

        # Для совместимости оставляем RECTOR, но заполняем пустой строкой
        replacements.update({
            '{{RECTOR}}': '',
            '{RECTOR}': '',
            'RECTOR': '',
            '{{PROVOST}}': '',
            '{PROVOST}': '',
            'PROVOST': '',
        })

        print("\n=== REPLACEMENTS TO APPLY ===")
        for key, value in replacements.items():
            if value:
                print(f"{key}: '{value}'")

        print("\nDEBUG: Applying replacements to document...")

        # Apply replacements
        self._replace_in_paragraphs(doc, replacements)
        self._replace_in_tables(doc, replacements)

        # 7. Create authors table
        self._create_authors_table(doc, authors_data)

        # 8. Create signature blocks
        self._create_signature_blocks(doc, authors_data)

        print("\n" + "=" * 80)
        print("DEBUG: Distribution agreement template filled successfully")
        print("=" * 80)

    def _create_authors_table(self, doc, authors_data):
        """Create table with authors information"""
        print(f"\n=== CREATING AUTHORS TABLE ===")

        if not authors_data or not authors_data.get('full_info_list'):
            print("No authors data for table")
            return

        authors_info = authors_data['full_info_list']
        print(f"Creating table for {len(authors_info)} authors")

        # Поиск и замена плейсхолдера TABLE
        found = False

        # Ищем в параграфах
        for paragraph in doc.paragraphs:
            if '{{TABLE}}' in paragraph.text or '{TABLE}' in paragraph.text or 'TABLE' in paragraph.text:
                print(f"Found table placeholder in paragraph")
                self._replace_paragraph_with_table(paragraph, authors_info)
                found = True
                break

        # Если не нашли в параграфах, ищем в таблицах
        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if '{{TABLE}}' in paragraph.text or '{TABLE}' in paragraph.text or 'TABLE' in paragraph.text:
                                print(f"Found table placeholder in table cell")
                                cell.text = ''
                                self._insert_table_in_cell(cell, authors_info)
                                found = True
                                break
                        if found:
                            break
                    if found:
                        break
                if found:
                    break

        if not found:
            print("WARNING: No table placeholder found in document")
            # Создаем таблицу в конце
            self._create_table_at_end(doc, authors_info)

    def _create_signature_blocks(self, doc, authors_data):
        """Create signature blocks for each author"""
        print(f"\n=== CREATING SIGNATURE BLOCKS ===")

        if not authors_data or not authors_data.get('full_info_list'):
            print("No authors data for signatures")
            return

        authors_info = authors_data['full_info_list']
        print(f"Creating signature blocks for {len(authors_info)} authors")

        # Ищем плейсхолдер SIGNATURE
        found = False

        # Ищем в параграфах
        for paragraph in doc.paragraphs:
            if '{{SIGNATURE}}' in paragraph.text or 'SIGNATURE' in paragraph.text or 'SIGNATURE' in paragraph.text:
                print(f"Found SIGNATURE placeholder in paragraph")
                self._replace_signature_placeholder(paragraph, authors_info)
                found = True
                break

        # Если не нашли в параграфах, ищем в таблицах
        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if '{{SIGNATURE}}' in paragraph.text or '{SIGNATURE}' in paragraph.text or 'SIGNATURE' in paragraph.text:
                                print(f"Found SIGNATURE placeholder in table cell")
                                self._create_signatures_in_cell(cell, authors_info)
                                found = True
                                break
                        if found:
                            break
                    if found:
                        break
                if found:
                    break

        if not found:
            print("WARNING: No SIGNATURE placeholder found in document")
            # Создаем подписи в конце документа
            self._create_signatures_at_end(doc, authors_info)

    def _replace_signature_placeholder(self, paragraph, authors_info):
        """Replace SIGNATURE placeholder with signature blocks"""
        # Получаем родительский элемент параграфа
        parent = paragraph._element.getparent()
        para_idx = parent.index(paragraph._element)

        # Удаляем параграф с плейсхолдером
        parent.remove(paragraph._element)

        # Создаем подписи на месте плейсхолдера
        for author_info in authors_info:
            full_name = author_info.get('full_name', '')

            # === ПЕРВАЯ СТРОКА ===
            # Линия 5.5 см (примерно 15 пробелов с подчеркиванием) → табуляция → линия 5 см (примерно 14 пробелов с подчеркиванием)
            first_line_para = OxmlElement('w:p')
            parent.insert(para_idx, first_line_para)
            para_idx += 1

            p1 = Document().add_paragraph()
            p1._p = first_line_para
            p1.paragraph_format.line_spacing = 1

            # Линия 5.5 см (примерно 15 символов)
            run1_1 = p1.add_run("_____________________________")  # ~5.5 см
            run1_1.font.name = 'Times New Roman'
            run1_1.font.size = Pt(12)

            # Табуляция
            run1_2 = p1.add_run("\t\t")
            run1_2.font.name = 'Times New Roman'
            run1_2.font.size = Pt(12)

            # Линия 5 см (примерно 14 символов)
            run1_3 = p1.add_run("__________________")  # ~5 см
            run1_3.font.name = 'Times New Roman'
            run1_3.font.size = Pt(12)

            # === ВТОРАЯ СТРОКА ===
            # Пробелы → ФИО → три табуляции → пробелы → (подпись)
            second_line_para = OxmlElement('w:p')
            parent.insert(para_idx, second_line_para)
            para_idx += 1

            p2 = Document().add_paragraph()
            p2._p = second_line_para
            p2.paragraph_format.line_spacing = 1

            # Пробелы для выравнивания (примерно 20 пробелов)
            run2_1 = p2.add_run("                    ")  # Пробелы для отступа

            # ФИО автора
            run2_2 = p2.add_run("(ФИО)")
            run2_2.font.name = 'Times New Roman'
            run2_2.font.size = Pt(10)

            # Три табуляции
            run2_3 = p2.add_run("\t\t\t")
            run2_3.font.name = 'Times New Roman'
            run2_3.font.size = Pt(10)

            # Пробелы (примерно 20 пробелов)
            run2_4 = p2.add_run("                       ")  # Пробелы для отступа

            # Слово "подпись" в скобках
            run2_5 = p2.add_run("(подпись)")
            run2_5.font.name = 'Times New Roman'
            run2_5.font.size = Pt(10)

            # Добавляем пустую строку между авторами (кроме последнего)
            if authors_info.index(author_info) < len(authors_info) - 1:
                empty_para = OxmlElement('w:p')
                parent.insert(para_idx, empty_para)
                para_idx += 1

                p_empty = Document().add_paragraph()
                p_empty._p = empty_para
                p_empty.paragraph_format.line_spacing = 1
                p_empty.add_run()

    def _create_signatures_in_cell(self, cell, authors_info):
        """Create signatures in a table cell"""
        # Очищаем ячейку
        cell.text = ''

        # Создаем подписи в ячейке
        for author_info in authors_info:
            full_name = author_info.get('full_name', '')

            # === ПЕРВАЯ СТРОКА ===
            p1 = cell.add_paragraph()
            p1.paragraph_format.space_after = Pt(6)
            p1.paragraph_format.line_spacing = 1

            # Линия 5.5 см
            run1_1 = p1.add_run("_____________________________")  # ~5.5 см
            run1_1.font.name = 'Times New Roman'
            run1_1.font.size = Pt(10)
            run1_1.underline = True

            # Табуляция
            run1_2 = p1.add_run("\t\t")
            run1_2.font.name = 'Times New Roman'
            run1_2.font.size = Pt(10)

            # Линия 5 см
            run1_3 = p1.add_run("__________________")  # ~5 см
            run1_3.font.name = 'Times New Roman'
            run1_3.font.size = Pt(10)
            run1_3.underline = True

            # === ВТОРАЯ СТРОКА ===
            p2 = cell.add_paragraph()
            p2.paragraph_format.space_after = Pt(6)
            p2.paragraph_format.line_spacing = 1

            # Пробелы для выравнивания
            run2_1 = p2.add_run("                    ")  # Пробелы для отступа

            # ФИО автора
            run2_2 = p2.add_run("(ФИО)")
            run2_2.font.name = 'Times New Roman'
            run2_2.font.size = Pt(10)

            # Три табуляции
            run2_3 = p2.add_run("\t\t\t")
            run2_3.font.name = 'Times New Roman'
            run2_3.font.size = Pt(10)

            # Пробелы для выравнивания
            run2_4 = p2.add_run("                       ")  # Пробелы для отступа

            # Слово "подпись" в скобках
            run2_5 = p2.add_run("(подпись)")
            run2_5.font.name = 'Times New Roman'
            run2_5.font.size = Pt(10)

            # Пустая строка между авторами
            if authors_info.index(author_info) < len(authors_info) - 1:
                p_empty = cell.add_paragraph()
                p_empty.paragraph_format.space_after = Pt(10)
                p_empty.paragraph_format.line_spacing = 1
                p_empty.add_run()

    def _create_signatures_at_end(self, doc, authors_info):
        """Create signatures at the end of document as fallback"""
        print("Creating signatures at end of document as fallback")

        # Добавляем пустую строку
        doc.add_paragraph()

        # Создаем подписи
        for author_info in authors_info:
            full_name = author_info.get('full_name', '')

            # === ПЕРВАЯ СТРОКА ===
            p1 = doc.add_paragraph()
            p1.paragraph_format.space_after = Pt(6)
            p1.paragraph_format.line_spacing = 1

            # Линия 5.5 см
            run1_1 = p1.add_run("_____________________________")  # ~5.5 см
            run1_1.font.name = 'Times New Roman'
            run1_1.font.size = Pt(10)
            run1_1.underline = True

            # Табуляция
            run1_2 = p1.add_run("\t\t")
            run1_2.font.name = 'Times New Roman'
            run1_2.font.size = Pt(10)

            # Линия 5 см
            run1_3 = p1.add_run("__________________")  # ~5 см
            run1_3.font.name = 'Times New Roman'
            run1_3.font.size = Pt(10)
            run1_3.underline = True

            # === ВТОРАЯ СТРОКА ===
            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(10)
            p2.paragraph_format.line_spacing = 1

            # Пробелы для выравнивания
            run2_1 = p2.add_run("                    ")  # Пробелы для отступа

            # ФИО автора
            run2_2 = p2.add_run("(ФИО)")
            run2_2.font.name = 'Times New Roman'
            run2_2.font.size = Pt(10)

            # Три табуляции
            run2_3 = p2.add_run("\t\t\t")
            run2_3.font.name = 'Times New Roman'
            run2_3.font.size = Pt(10)

            # Пробелы для выравнивания
            run2_4 = p2.add_run("                       ")  # Пробелы для отступа

            # Слово "подпись" в скобках
            run2_5 = p2.add_run("(подпись)")
            run2_5.font.name = 'Times New Roman'
            run2_5.font.size = Pt(10)

            # Пустая строка между авторами
            if authors_info.index(author_info) < len(authors_info) - 1:
                doc.add_paragraph()

    def _replace_paragraph_with_table(self, paragraph, authors_info):
        """Replace paragraph with table"""
        # Получаем родительский элемент параграфа
        parent = paragraph._element.getparent()

        # Создаем таблицу
        table = self._create_formatted_table(authors_info)

        # Вставляем таблицу после параграфа
        parent.insert(parent.index(paragraph._element) + 1, table._element)

        # Удаляем параграф с плейсхолдером
        parent.remove(paragraph._element)

        print("Paragraph replaced with table")

    def _insert_table_in_cell(self, cell, authors_info):
        """Insert table in a table cell"""
        # Создаем таблицу
        table = self._create_formatted_table(authors_info)

        # Очищаем ячейку и добавляем таблицу
        cell._element.clear()
        cell._element.append(table._element)

        print("Table inserted in cell")

    def _create_formatted_table(self, authors_info):
        """Create formatted table with borders and Times New Roman font"""
        # Создаем таблицу
        table = Document().add_table(rows=len(authors_info) + 1, cols=4)

        # Убираем стиль таблицы
        try:
            table.style = None
        except:
            pass

        # Устанавливаем выравнивание по центру
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Устанавливаем ширину колонок
        widths = [Inches(0.6), Inches(3.5), Inches(3.0), Inches(1.5)]
        for i, width in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = width

        # Применяем границы ко всей таблице
        self._apply_table_borders(table)

        # Заполняем заголовок таблицы
        header_cells = table.rows[0].cells
        headers = [
            "№ п/п",
            "Ф.И.О. полностью",
            "Должность, место работы,\nдля обучающегося – уровень образования\n(аспирант, магистр, специалист, бакалавр группа)",
            "Показатель\nтворческого участия, %"
        ]

        for idx, header in enumerate(headers):
            # Очищаем ячейку
            header_cells[idx].text = ''

            # Добавляем текст с форматированием
            p = header_cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = p.add_run(header)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.bold = True

        # Заполняем строки с данными авторов
        for row_idx, author_info in enumerate(authors_info):
            data_row = table.rows[row_idx + 1].cells

            # Колонка 1: Порядковый номер
            data_row[0].text = str(row_idx + 1)
            for paragraph in data_row[0].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

            # Колонка 2: ФИО полностью
            data_row[1].text = author_info.get('full_name', '')
            for paragraph in data_row[1].paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

            # Колонка 3: Статус/должность/образование в родительном падеже
            # Используем поле 'status_for_table_genitive'
            status_genitive = author_info.get('status_for_table_genitive', '')

            data_row[2].text = status_genitive
            for paragraph in data_row[2].paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

            # Колонка 4: Процент участия
            percentage = author_info.get('percentage', '')
            # Если процент есть, но без знака %, добавляем его
            if percentage and '%' not in str(percentage):
                percentage = f"{percentage}%"

            data_row[3].text = percentage if percentage else ""
            for paragraph in data_row[3].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12)

            print(f"  Row {row_idx + 1}: {author_info.get('full_name')} - {status_genitive} - {percentage}")

        return table

    def _apply_table_borders(self, table):
        """Apply borders to all cells in the table"""
        tbl = table._element
        tblPr = tbl.tblPr

        # Создаем границы для таблицы
        tblBorders = OxmlElement('w:tblBorders')

        # Добавляем все границы
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)

        # Устанавливаем границы таблицы
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.append(tblPr)
        tblPr.append(tblBorders)

        # Применяем границы ко всем ячейкам
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()

                # Убираем существующие границы
                existing_borders = tcPr.findall(qn('w:tcBorders'))
                for border in existing_borders:
                    tcPr.remove(border)

                # Добавляем новые границы
                tcBorders = OxmlElement('w:tcBorders')
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '4')
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), '000000')
                    tcBorders.append(border)
                tcPr.append(tcBorders)

    def _create_table_at_end(self, doc, authors_info):
        """Create table at the end of document as fallback"""
        print("Creating table at end of document as fallback")

        # Добавляем пустой параграф
        doc.add_paragraph()

        # Создаем таблицу
        table = self._create_formatted_table(authors_info)

        # Добавляем таблицу в документ
        doc._element.body.append(table._element)

    def _extract_nir_info(self, application):
        """Extract NIR information from application"""
        print(f"\nDEBUG _extract_nir_info: Processing application for NIR info")

        if not application or not isinstance(application, tuple):
            print("  Invalid application data")
            return "Научно-исследовательская работа"

        # Показываем все поля заявки для отладки
        print(f"  Application tuple length: {len(application)}")
        for i, field in enumerate(application):
            field_type = type(field).__name__
            field_value = str(field)
            print(f"    [{i}] {field_type}: '{field_value}'")

        # Из структуры DAO видно, что запрос возвращает поля:
        # 0: id, 1: title, 2: product_id, 3: user_id, 4: commission_id,
        # 5: status_id, 6: created_date, 7: nir, 8: product_name, 9: username,
        # 10: commission_name, 11: status_name, 12: status_id (повторно?)

        # Судя по логу, в индексе 8 находится product_name ('Латэшка')
        # НИР должна быть в индексе 7

        # Проверяем индекс 7 - это должно быть поле nir
        if len(application) > 7 and application[7]:
            nir_value = str(application[7]).strip()
            print(f"  Found NIR at index 7: '{nir_value}'")

            if nir_value and nir_value.lower() != 'admin' and nir_value != 'None':
                return nir_value

        # Если в индексе 7 нет NIR, проверяем индекс 6 (created_date) или 5 (status_id)
        # Но это не NIR, поэтому возвращаем название продукта
        print(f"  No NIR found at index 7")

        # Возвращаем название продукта как тему НИР
        if len(application) > 8 and application[8]:
            product_name = str(application[8]).strip()
            if product_name and product_name.lower() != 'admin':
                return f"Научно-исследовательская работа по теме: {product_name}"

        return "Научно-исследовательская работа"

    def _extract_product_name(self, application):
        """Extract product name from application"""
        if not application or not isinstance(application, tuple):
            print("WARNING: Invalid application data in _extract_product_name")
            return ''

        # Check indices in priority order
        indices_to_try = [8, 7, 1]

        for idx in indices_to_try:
            if len(application) > idx and application[idx]:
                name = str(application[idx]).strip()
                if name and name.lower() != 'admin' and not name.isdigit():
                    return name

        # If no proper name found, return first non-empty
        for idx in indices_to_try:
            if len(application) > idx and application[idx]:
                return str(application[idx]).strip()

        print("WARNING: No product name found in application")
        return ''

    def _extract_authors_from_list(self, authors_list):
        """Extract authors data from list of author tuples - ИСПРАВЛЕННЫЙ МЕТОД"""
        print(f"\nDEBUG _extract_authors_from_list: Processing {len(authors_list)} authors")

        result = {
            'full_info_list': [],
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
                # Показываем все поля для отладки
                for i, field in enumerate(author):
                    field_type = type(field).__name__
                    field_value = str(field)
                    print(f"      [{i}] {field_type}: '{field_value}'")

                # Инициализация переменных
                surname = ''
                name = ''
                patronymic = ''
                position = ''
                department = ''
                status_for_table = ''
                status_genitive = ''
                percentage = ''
                birth_date = ''
                citizenship = ''
                address = ''
                is_student = False
                education_level = ''
                education_level_id = ''
                role_name = ''

                # Извлекаем основные поля согласно структуре в логе:
                if len(author) >= 19:  # Убедимся, что у нас достаточно полей
                    # ФИО
                    surname = str(author[7]).strip() if author[7] else ''
                    name = str(author[8]).strip() if author[8] else ''
                    patronymic = str(author[9]).strip() if author[9] else ''

                    # Личная информация
                    birth_date = author[10] if len(author) > 10 else ''
                    citizenship = str(author[11]).strip() if len(author) > 11 and author[11] else ''
                    address = str(author[12]).strip() if len(author) > 12 and author[12] else ''

                    # Статус и образование
                    is_student = bool(author[13]) if len(author) > 13 and author[13] is not None else False

                    # Образование
                    education_level_id = str(author[15]).strip() if len(author) > 15 and author[15] else ''
                    education_level = str(author[16]).strip() if len(author) > 16 and author[16] else ''

                    # Должность/роль
                    position = str(author[17]).strip() if len(author) > 17 and author[17] else ''

                    # Подразделение (department) - индекс 18
                    department = str(author[18]).strip() if len(author) > 18 and author[18] else ''

                # Процент участия
                if len(author) > 3 and author[3] is not None:
                    try:
                        if hasattr(author[3], 'to_eng_string'):
                            percentage_str = author[3].to_eng_string()
                        else:
                            percentage_str = str(author[3])

                        perc_float = float(percentage_str)

                        # Если число меньше 1 (например, 0.10), умножаем на 100
                        if perc_float < 1:
                            perc_float = perc_float * 100

                        # Форматируем без лишних нулей
                        if perc_float.is_integer():
                            percentage = str(int(perc_float))
                        else:
                            percentage = f"{perc_float:.2f}".rstrip('0').rstrip('.')
                    except Exception as e:
                        print(f"    Error processing percentage '{author[3]}': {e}")
                        percentage = str(author[3]) if author[3] else ''

                # Формируем строку для таблицы
                # Правила:
                # 1. Если is_student = True, то показываем уровень образования
                # 2. Если is_student = False, то показываем должность
                # 3. В любом случае добавляем подразделение, если оно есть
                # 4. Только ПОДРАЗДЕЛЕНИЕ переводим в родительный падеж

                if is_student:
                    # Для студента используем уровень образования
                    if education_level:
                        if department:
                            # Преобразуем только подразделение в родительный падеж
                            department_genitive = self._to_genitive_department(department)
                            status_for_table = f"{education_level} {department_genitive}"
                            status_genitive = f"{education_level} {department_genitive}"
                        else:
                            status_for_table = education_level
                            status_genitive = education_level
                    else:
                        # Если уровень образования не указан
                        if department:
                            department_genitive = self._to_genitive_department(department)
                            status_for_table = f"студент {department_genitive}"
                            status_genitive = f"студент {department_genitive}"
                        else:
                            status_for_table = "студент"
                            status_genitive = "студент"
                else:
                    # Для не-студента используем должность
                    if position:
                        if department:
                            # Преобразуем только подразделение в родительный падеж
                            department_genitive = self._to_genitive_department(department)
                            status_for_table = f"{position} {department_genitive}"
                            status_genitive = f"{position} {department_genitive}"
                        else:
                            status_for_table = position
                            status_genitive = position
                    else:
                        # Если должность не указана
                        if department:
                            # Преобразуем только подразделение в родительный падеж
                            department_genitive = self._to_genitive_department(department)
                            status_for_table = department_genitive
                            status_genitive = department_genitive
                        else:
                            status_for_table = ""
                            status_genitive = ""

                print(f"    Final - Name: {surname} {name} {patronymic}")
                print(f"    Final - Is student: {is_student}")
                print(f"    Final - Education level: '{education_level}'")
                print(f"    Final - Position: '{position}'")
                print(f"    Final - Department: '{department}'")
                print(f"    Final - Status for table: '{status_for_table}'")
                print(f"    Final - Percentage: '{percentage}'")

                if surname and name:
                    full_name = f"{surname} {name} {patronymic}".strip()

                    # Форматируем процент для отображения
                    display_percentage = ''
                    if percentage:
                        try:
                            perc_num = float(percentage)
                            if perc_num.is_integer():
                                display_percentage = str(int(perc_num))
                            else:
                                formatted = f"{perc_num:.2f}"
                                display_percentage = formatted.rstrip('0').rstrip('.')
                            display_percentage += '%'
                        except Exception as e:
                            print(f"    Error formatting percentage '{percentage}': {e}")
                            display_percentage = str(percentage)

                    # Создаем объект с полной информацией об авторе
                    author_info = {
                        'full_name': full_name,
                        'surname': surname,
                        'name': name,
                        'patronymic': patronymic,
                        'position': position,
                        'department': department,
                        'status_for_table': status_for_table,
                        'status_for_table_genitive': status_genitive,
                        'is_student': is_student,
                        'education_level': education_level,
                        'education_level_id': education_level_id,
                        'percentage': display_percentage,
                        'birth_date': birth_date,
                        'citizenship': citizenship,
                        'address': address
                    }

                    result['full_info_list'].append(author_info)
                    result['full_names_list'].append(full_name)
                    print(f"    ✓ Added author: {full_name}")

            elif isinstance(author, dict):
                # Если автор уже в виде словаря
                full_name = author.get('full_name', '')
                if full_name:
                    result['full_info_list'].append(author)
                    result['full_names_list'].append(full_name)
                    print(f"    ✓ Added author from dict: {full_name}")

            elif isinstance(author, str):
                full_name = author.strip()
                if full_name:
                    author_info = {
                        'full_name': full_name,
                        'position': '',
                        'department': '',
                        'status_for_table': '',
                        'status_for_table_genitive': '',
                        'is_student': False,
                        'education_level': '',
                        'percentage': ''
                    }
                    result['full_info_list'].append(author_info)
                    result['full_names_list'].append(full_name)
                    print(f"    ✓ String author: {full_name}")

        # Форматируем список авторов для текстовых полей
        if result['full_names_list']:
            if len(result['full_names_list']) == 1:
                result['authors_list_formatted'] = result['full_names_list'][0]
            else:
                authors_str = ', '.join(result['full_names_list'][:-1])
                result['authors_list_formatted'] = f"{authors_str} и {result['full_names_list'][-1]}"

            # Родительный падеж для подписей
            genitive_list = []
            for author_info in result['full_info_list']:
                author_name = author_info['full_name']
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
                else:
                    genitive_list.append(author_name)

            if len(genitive_list) == 1:
                result['authors_genitive_formatted'] = genitive_list[0]
            else:
                genitive_str = ', '.join(genitive_list[:-1])
                result['authors_genitive_formatted'] = f"{genitive_str} и {genitive_list[-1]}"

        print(f"\n  Final: {len(result['full_info_list'])} authors with full info")
        return result

    def _to_genitive_department(self, department):
        """Convert department to genitive case"""
        if not department:
            return ''

        dept_lower = department.lower().strip()

        # Правила для русских названий подразделений
        if dept_lower.endswith('кафедра'):
            return department[:-6] + 'кафедры'
        elif dept_lower.endswith('отдел'):
            return department[:-5] + 'отдела'
        elif dept_lower.endswith('лаборатория'):
            return department[:-11] + 'лаборатории'
        elif dept_lower.endswith('факультет'):
            return department[:-9] + 'факультета'
        elif dept_lower.endswith('институт'):
            return department[:-8] + 'института'
        elif dept_lower.endswith('центр'):
            return department[:-5] + 'центра'
        elif dept_lower.endswith('управление'):
            return department[:-10] + 'управления'
        elif dept_lower.endswith('департамент'):
            return department[:-11] + 'департамента'
        elif dept_lower.endswith('служба'):
            return department[:-6] + 'службы'
        elif dept_lower.endswith('бюро'):
            return department[:-4] + 'бюро'
        elif dept_lower.endswith('сектор'):
            return department[:-6] + 'сектора'
        elif dept_lower.endswith('группа'):
            return department[:-6] + 'группы'
        elif dept_lower.endswith('команда'):
            return department[:-7] + 'команды'

        # Общие правила для русских слов
        elif dept_lower.endswith('ая'):
            return department[:-2] + 'ой'  # научная -> научной
        elif dept_lower.endswith('яя'):
            return department[:-2] + 'ей'  # ведущая -> ведущей
        elif dept_lower.endswith('ий'):
            return department[:-2] + 'его'  # научный -> научного
        elif dept_lower.endswith('ый'):
            return department[:-2] + 'ого'  # учебный -> учебного
        elif dept_lower.endswith('ой'):
            return department[:-2] + 'ого'  # рабочий -> рабочего
        elif dept_lower.endswith('а'):
            return department[:-1] + 'ы'  # кафедра -> кафедры
        elif dept_lower.endswith('я'):
            return department[:-1] + 'и'  # лаборатория -> лаборатории
        elif dept_lower.endswith('о'):
            return department[:-1] + 'а'  # бюро -> бюро (не меняется)
        elif dept_lower.endswith('е'):
            return department[:-1] + 'я'  # управление -> управления
        elif dept_lower.endswith('ь'):
            return department[:-1] + 'я'  # служба -> службы
        else:
            # По умолчанию добавляем 'а'
            return department + 'а'

    def _to_genitive_surname(self, surname):
        """Convert surname to genitive case (simplified)"""
        if not surname:
            return ''

        surname_lower = surname.strip().lower()

        # Russian language rules
        if surname_lower.endswith('ова') or surname_lower.endswith('ева') or surname_lower.endswith('ина'):
            return surname[:-1] + 'ой'
        elif surname_lower.endswith('ая'):
            return surname[:-2] + 'ой'
        elif surname_lower.endswith('ская'):
            return surname[:-4] + 'ской'
        elif surname_lower.endswith('цкая'):
            return surname[:-4] + 'цкой'
        elif surname_lower.endswith('кая'):
            return surname[:-3] + 'кой'
        elif surname_lower.endswith('га') or surname_lower.endswith('ка'):
            return surname[:-1] + 'и'
        elif surname_lower.endswith('я'):
            return surname[:-1] + 'и'
        elif surname_lower.endswith('ов') or surname_lower.endswith('ев') or surname_lower.endswith('ин'):
            return surname + 'а'
        elif surname_lower.endswith('ий') or surname_lower.endswith('ый'):
            return surname[:-2] + 'ого'
        elif surname_lower.endswith('ой'):
            return surname[:-2] + 'ого'
        else:
            return surname + 'а'

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
