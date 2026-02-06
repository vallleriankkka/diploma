import docx
from flask import Flask, render_template, request, redirect, url_for, jsonify, flash, session, send_file, current_app, \
    Response
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from functools import wraps
import tempfile
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from io import BytesIO
import os
import zipfile
import sys

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='ignore')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='ignore')

# Импорт собственных классов
from database import Database
from dao.user_dao import UserDAO
from dao.author_dao import AuthorDAO
from dao.product_dao import ProductDAO
from dao.application_dao import ApplicationDAO
from dao.expert_dao import ExpertDAO
from dao.commission_dao import CommissionDAO
from dao.commission_expert_dao import CommissionExpertDAO
from dao.setting_dao import SettingDAO
from service.author_service import AuthorService
from documents.abstract_generator import AbstractGenerator
from documents.sogl_generator import SoglGenerator
from documents.consent_generator import ConsentGenerator
from documents.distributionagreement_generator import DistributionAgreementGenerator
from documents.conclusion_generator import ConclusionGenerator
from documents.notification_generator import RIDNotificationGenerator
from documents.assignment_generator import AssignmentAgreementGenerator

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-123'
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Инициализация базы данных и DAO
db = Database.get_instance()
user_dao = UserDAO(db)
author_dao = AuthorDAO(db)
product_dao = ProductDAO(db)
application_dao = ApplicationDAO(db)
expert_dao = ExpertDAO(db)
commission_dao = CommissionDAO(db)
setting_dao = SettingDAO(db)
commission_expert_dao = CommissionExpertDAO(db)
# Инициализация сервисов
author_service = AuthorService(author_dao)


# Декораторы для проверки прав доступа
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Для доступа к этой странице необходимо войти в систему', 'error')
            return redirect(url_for('login'))
        return f(*args, **kwargs)

    return decorated_function


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Для доступа к этой странице необходимо войти в систему', 'error')
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('У вас недостаточно прав для доступа к этой странице', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)

    return decorated_function


# Jinja2 фильтр для безопасного форматирования даты
@app.template_filter('safe_strftime')
def safe_strftime(value, format='%d.%m.%Y %H:%M'):
    if value is None:
        return ''

    try:
        # Если это уже datetime объект
        if hasattr(value, 'strftime'):
            return value.strftime(format)

        # Если это строка, пробуем преобразовать
        if isinstance(value, str):
            # Пробуем разные форматы
            formats_to_try = [
                '%Y-%m-%d',
                '%Y-%m-%d %H:%M:%S',
                '%d.%m.%Y',
                '%d.%m.%Y %H:%M:%S',
                '%Y-%m-%dT%H:%M:%S',
                '%Y-%m-%dT%H:%M:%S.%fZ'
            ]

            for fmt in formats_to_try:
                try:
                    dt_obj = datetime.strptime(value, fmt)
                    return dt_obj.strftime(format)
                except ValueError:
                    continue

            # Если не удалось преобразовать, возвращаем как есть
            return value

        # Для других типов просто преобразуем в строку
        return str(value)

    except Exception:
        # В случае любой ошибки возвращаем исходное значение как строку
        return str(value)


# Регистрируем фильтр
app.jinja_env.filters['safe_strftime'] = safe_strftime


# ========== АДМИН ПАНЕЛЬ - НАСТРОЙКИ СУЩНОСТЕЙ ==========
@app.route('/settings', methods=['GET'])
@login_required
def settings():
    """Страница настроек"""
    if session['role'] != 'admin':
        flash('Доступ запрещен', 'error')
        return redirect(url_for('index'))

    try:
        # Получаем текущие настройки
        current_settings_tuple = setting_dao.get_current()

        # Преобразуем tuple в словарь для удобства использования в шаблоне
        current_settings = {}
        if current_settings_tuple:
            current_settings = {
                'id': current_settings_tuple[0],
                'assignee_id': current_settings_tuple[1],
                'power_of_attorney_id': current_settings_tuple[2],
                'sum_id': current_settings_tuple[3],
                'provost_id': current_settings_tuple[4],
                'assignee_name': current_settings_tuple[5],
                'assignee_post': current_settings_tuple[6],
                'poa_number': current_settings_tuple[7],
                'poa_date': current_settings_tuple[8],
                'sum_value': current_settings_tuple[9],
                'provost_name': current_settings_tuple[10],
                'provost_post': current_settings_tuple[11]
            }

        # Загружаем опции для выпадающих списков
        assignee_options = setting_dao.get_reference_options('assignee')
        power_of_attorney_options = setting_dao.get_reference_options('power_of_attorney')
        sum_options = setting_dao.get_reference_options('sum')
        provost_options = setting_dao.get_reference_options('provost')

        # Загружаем справочники для вкладки управления
        programming_languages = product_dao.get_all_programming_languages()
        os_types = product_dao.get_all_os_types()
        computer_types = product_dao.get_all_computer_types()

        # Загружаем роли авторов
        author_roles = application_dao.get_all_author_roles()

        return render_template('settings/index.html',
                               current_settings=current_settings,
                               assignee_options=assignee_options,
                               power_of_attorney_options=power_of_attorney_options,
                               sum_options=sum_options,
                               provost_options=provost_options,
                               programming_languages=programming_languages,
                               os_types=os_types,
                               computer_types=computer_types,
                               author_roles=author_roles)
    except Exception as e:
        flash(f'Ошибка загрузки настроек: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/admin/settings/save', methods=['POST'])
@login_required
@admin_required
def save_settings():
    """Сохранение настроек"""
    try:
        # Подготавливаем данные для сохранения
        setting_data = {
            'assignee_id': request.form.get('assignee_id') or None,
            'power_of_attorney_id': request.form.get('power_of_attorney_id') or None,
            'sum_id': request.form.get('sum_id') or None,
            'provost_id': request.form.get('provost_id') or None
        }

        # Сохраняем настройки
        setting_id = setting_dao.save(setting_data)

        flash('Настройки успешно сохранены!', 'success')
        return redirect(url_for('settings_main'))

    except Exception as e:
        flash(f'Ошибка при сохранении настроек: {str(e)}', 'error')
        return redirect(url_for('settings_main'))


@app.route('/admin/settings/update_single', methods=['POST'])
@login_required
@admin_required
def update_single_setting():
    """Обновление одной настройки"""
    try:
        setting_type = request.form.get('setting_type')
        reference_id = request.form.get('reference_id')

        if not setting_type:
            return jsonify({'success': False, 'error': 'Не указан тип настройки'})

        # Преобразуем reference_id в int, если он есть
        ref_id = int(reference_id) if reference_id else None

        # Сопоставляем тип настройки с полем в таблице
        field_mapping = {
            'assignee': 'assignee_id',
            'power_of_attorney': 'power_of_attorney_id',
            'sum': 'sum_id',
            'provost': 'provost_id'
        }

        field_name = field_mapping.get(setting_type)
        if not field_name:
            return jsonify({'success': False, 'error': 'Неверный тип настройки'})

        # Обновляем настройку
        success = setting_dao.update_single(field_name, ref_id)

        if success:
            return jsonify({'success': True, 'message': 'Настройка обновлена'})
        else:
            return jsonify({'success': False, 'error': 'Не удалось обновить настройку'})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ===== Роуты для создания новых записей в справочниках =====

@app.route('/admin/settings/create_assignee', methods=['POST'])
@login_required
@admin_required
def create_assignee():
    """Создание нового доверенного лица"""
    try:
        full_name = request.form.get('full_name', '').strip()
        post = request.form.get('post', '').strip()

        if not full_name:
            return jsonify({'success': False, 'error': 'ФИО не может быть пустым'})

        # Проверяем, не существует ли уже такое ФИО
        if setting_dao.check_exists('assignee', 'full_name', full_name):
            return jsonify({'success': False, 'error': 'Доверенное лицо с таким ФИО уже существует'})

        # Создаем новую запись
        assignee_id = setting_dao.create_assignee(full_name, post or None)

        # Обновляем текущие настройки
        setting_dao.update_single('assignee_id', assignee_id)

        display_name = f"{full_name}" + (f" ({post})" if post else "")

        return jsonify({
            'success': True,
            'message': 'Доверенное лицо успешно создано',
            'id': assignee_id,
            'name': full_name,
            'display_name': display_name,
            'post': post
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/create_power_of_attorney', methods=['POST'])
@login_required
@admin_required
def create_power_of_attorney():
    """Создание новой доверенности"""
    try:
        number = request.form.get('number', '').strip()
        date_of_receipt = request.form.get('date_of_receipt', '')

        if not number:
            return jsonify({'success': False, 'error': 'Номер доверенности не может быть пустым'})

        if not date_of_receipt:
            return jsonify({'success': False, 'error': 'Дата получения не может быть пустой'})

        # Проверяем, не существует ли уже такой доверенности
        if setting_dao.check_exists('power_of_attorney', 'number', number):
            return jsonify({'success': False, 'error': 'Доверенность с таким номером уже существует'})

        # Создаем новую запись
        poa_id = setting_dao.create_power_of_attorney(number, date_of_receipt)

        # Обновляем текущие настройки
        setting_dao.update_single('power_of_attorney_id', poa_id)

        from datetime import datetime
        date_obj = datetime.strptime(date_of_receipt, '%Y-%m-%d')
        date_str = date_obj.strftime('%d.%m.%Y')

        return jsonify({
            'success': True,
            'message': 'Доверенность успешно создана',
            'id': poa_id,
            'name': number,
            'display_name': f"{number} ({date_str})",
            'date': date_of_receipt
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/create_sum', methods=['POST'])
@login_required
@admin_required
def create_sum():
    """Создание новой суммы"""
    try:
        value_str = request.form.get('value', '').strip()

        if not value_str:
            return jsonify({'success': False, 'error': 'Сумма не может быть пустой'})

        try:
            value = float(value_str)
            if value <= 0:
                return jsonify({'success': False, 'error': 'Сумма должна быть положительным числом'})
        except ValueError:
            return jsonify({'success': False, 'error': 'Введите корректное числовое значение'})

        # Проверяем, не существует ли уже такая сумма
        if setting_dao.check_exists('sum', 'value', value):
            return jsonify({'success': False, 'error': 'Сумма с таким значением уже существует'})

        # Создаем новую запись
        sum_id = setting_dao.create_sum(value)

        # Обновляем текущие настройки
        setting_dao.update_single('sum_id', sum_id)

        return jsonify({
            'success': True,
            'message': 'Сумма успешно создана',
            'id': sum_id,
            'name': str(value),
            'display_name': str(value)
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/create_provost', methods=['POST'])
@login_required
@admin_required
def create_provost():
    """Создание нового проректора"""
    try:
        full_name = request.form.get('full_name', '').strip()
        post = request.form.get('post', '').strip()

        if not full_name:
            return jsonify({'success': False, 'error': 'ФИО не может быть пустым'})

        # Проверяем, не существует ли уже такое ФИО
        if setting_dao.check_exists('provost', 'full_name', full_name):
            return jsonify({'success': False, 'error': 'Проректор с таким ФИО уже существует'})

        # Создаем новую запись
        provost_id = setting_dao.create_provost(full_name, post or None)

        # Обновляем текущие настройки
        setting_dao.update_single('provost_id', provost_id)

        display_name = f"{full_name}" + (f" ({post})" if post else "")

        return jsonify({
            'success': True,
            'message': 'Проректор успешно создан',
            'id': provost_id,
            'name': full_name,
            'display_name': display_name,
            'post': post
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ========== НАСТРОЙКИ СПРАВОЧНИКОВ ==========

@app.route('/admin/settings/create_reference', methods=['POST'])
@login_required
@admin_required
def create_reference():
    """Создание новой записи в справочнике"""
    try:
        data = request.get_json()
        reference_type = data.get('reference_type')
        name = data.get('name', '').strip()

        if not reference_type:
            return jsonify({'success': False, 'error': 'Не указан тип справочника'})

        if not name:
            return jsonify({'success': False, 'error': 'Название не может быть пустым'})

        # Проверяем, не существует ли уже такое название
        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            return jsonify({'success': False, 'error': 'Неизвестный тип справочника'})

        # Проверка существования
        with product_dao.db.get_cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {table_name} WHERE name = %s", (name,))
            if cur.fetchone()[0] > 0:
                return jsonify({'success': False, 'error': 'Такое название уже существует'})

        # Создаем новую запись
        reference = product_dao.create_reference(reference_type, name)

        return jsonify({
            'success': True,
            'message': 'Запись успешно создана',
            'reference': reference
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/update_reference', methods=['POST'])
@login_required
@admin_required
def update_reference():
    """Обновление записи в справочнике"""
    try:
        data = request.get_json()
        reference_type = data.get('reference_type')
        ref_id = data.get('id')
        name = data.get('name', '').strip()

        if not reference_type:
            return jsonify({'success': False, 'error': 'Не указан тип справочника'})

        if not ref_id:
            return jsonify({'success': False, 'error': 'Не указан ID записи'})

        if not name:
            return jsonify({'success': False, 'error': 'Название не может быть пустым'})

        # Проверяем, не существует ли уже такое название (кроме текущей записи)
        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            return jsonify({'success': False, 'error': 'Неизвестный тип справочника'})

        # Проверка существования другого элемента с таким же названием
        with product_dao.db.get_cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {table_name} WHERE name = %s AND id != %s",
                        (name, ref_id))
            if cur.fetchone()[0] > 0:
                return jsonify({'success': False, 'error': 'Такое название уже существует'})

        # Обновляем запись
        success = product_dao.update_reference(reference_type, ref_id, name)

        if success:
            return jsonify({
                'success': True,
                'message': 'Запись успешно обновлена',
                'reference': {'id': ref_id, 'name': name}
            })
        else:
            return jsonify({'success': False, 'error': 'Запись не найдена'})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/delete_reference', methods=['POST'])
@login_required
@admin_required
def delete_reference():
    """Удаление записи из справочника"""
    try:
        data = request.get_json()
        reference_type = data.get('reference_type')
        ref_id = data.get('id')

        if not reference_type:
            return jsonify({'success': False, 'error': 'Не указан тип справочника'})

        if not ref_id:
            return jsonify({'success': False, 'error': 'Не указан ID записи'})

        # Удаляем запись
        success = product_dao.delete_reference(reference_type, ref_id)

        if success:
            return jsonify({
                'success': True,
                'message': 'Запись успешно удалена'
            })
        else:
            return jsonify({'success': False, 'error': 'Запись не найдена'})

    except ValueError as e:
        # Обрабатываем ошибку "используется в продуктах"
        return jsonify({'success': False, 'error': str(e)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/settings/create_author_role', methods=['POST'])
@login_required
@admin_required
def create_author_role():
    """Создание новой роли автора"""
    try:
        data = request.get_json()
        name = data.get('name', '').strip()

        if not name:
            return jsonify({'success': False, 'error': 'Название роли не может быть пустым'})

        # Проверяем, не существует ли уже такая роль
        with application_dao.db.get_cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM expert_roles WHERE name = %s", (name,))
            if cur.fetchone()[0] > 0:
                return jsonify({'success': False, 'error': 'Роль с таким названием уже существует'})

        # Создаем новую запись
        with application_dao.db.get_cursor() as cur:
            cur.execute("""
                    INSERT INTO expert_roles (name) 
                    VALUES (%s) 
                    RETURNING id, name
                """, (name,))
            result = cur.fetchone()
            role = {'id': result[0], 'name': result[1]}

        return jsonify({
            'success': True,
            'message': 'Роль автора успешно создана',
            'reference': role
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ========== АУТЕНТИФИКАЦИЯ ==========
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        confirm_password = request.form['confirm_password']

        if password != confirm_password:
            flash('Пароли не совпадают', 'error')
            return render_template('auth/register.html', username=username)

        success, result = user_dao.create_user(username, password)

        if success:
            flash('Регистрация успешна! Теперь войдите в систему', 'success')
            return redirect(url_for('login'))
        else:
            # Теперь result - это словарь
            error_msg = result.get('error', 'Неизвестная ошибка')
            if 'уже существует' in error_msg:
                flash('Пользователь с таким именем уже существует', 'error')
            else:
                flash(f'Ошибка при регистрации: {error_msg}', 'error')
            return render_template('auth/register.html', username=username)

    return render_template('auth/register.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        user = user_dao.get_user_by_username(username)

        if user and user_dao.verify_password(user, password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            flash(f'Добро пожаловать, {user["username"]}!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Неверное имя пользователя или пароль', 'error')
            return render_template('auth/login.html', username=username)

    return render_template('auth/login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('Вы вышли из системы', 'success')
    return redirect(url_for('index'))


# ========== ГЛАВНАЯ СТРАНИЦА ==========
@app.route('/')
def index():
    return render_template('index.html')


# ========== АВТОРЫ ==========
@app.route('/authors')
@login_required
def authors_list():
    if session['role'] == 'admin':
        authors = author_dao.get_all(with_education_info=True)  # Админ видит всех
    else:
        authors = author_dao.get_all(session['user_id'], with_education_info=True)  # Пользователь только своих

    # Преобразуем даты во всех авторах
    processed_authors = []
    for author in authors:
        author_list = list(author)
        # Преобразуем строки дат в объекты date
        for i in [4, 8]:  # birth_date, passport_issue_date
            if author_list[i] and isinstance(author_list[i], str):
                try:
                    author_list[i] = datetime.strptime(author_list[i], '%Y-%m-%d').date()
                except ValueError:
                    author_list[i] = None
        processed_authors.append(author_list)

    return render_template('authors/list.html', authors=processed_authors)


@app.route('/author/create', methods=['GET', 'POST'])
@login_required
def author_create():
    # Получаем уровни образования один раз
    education_levels = get_education_levels()

    if request.method == 'POST':
        try:
            # Подготавливаем дату выдачи паспорта (обязательное поле)
            passport_issue_date_str = request.form.get('passport_issue_date')
            if not passport_issue_date_str:
                flash('Дата выдачи паспорта обязательна для заполнения', 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)  # Передаем все данные формы

            passport_issue_date = datetime.strptime(passport_issue_date_str, '%Y-%m-%d').date()

            # Обработка чекбокса студента - получаем как строку
            is_student = request.form.get('is_student') == 'true'
            education_level_id = request.form.get('education_level_id')

            # Если студент, проверяем уровень образования
            if is_student and not education_level_id:
                flash('Для студента необходимо указать уровень образования', 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)

            # Подготавливаем данные для валидации
            birth_date_str = request.form.get('birth_date')
            if not birth_date_str:
                flash('Дата рождения обязательна', 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)

            birth_date = datetime.strptime(birth_date_str, '%Y-%m-%d').date()

            form_data = {
                'surname': request.form['surname'],
                'name': request.form['name'],
                'patronymic': request.form.get('patronymic', ''),
                'birth_date': birth_date,
                'address': request.form['address'],
                'department': request.form.get('department', ''),
                'passport': request.form['passport'],
                'passport_issue_date': passport_issue_date,
                'passport_issued_by': request.form.get('passport_issued_by', ''),
                'passport_issued_code': request.form.get('passport_issued_code', ''),
                'snils': request.form['snils'],
                'inn': request.form['inn'],
                'citizenship': request.form.get('citizenship', 'Российская Федерация'),
                'position': request.form.get('position', ''),
                'is_student': is_student,
                'user_id': session['user_id'],
                'education_level_id': int(education_level_id) if education_level_id else None
            }

            # Валидация
            errors = author_service.validate_author(form_data)
            if errors:
                for error in errors:
                    flash(error, 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)

            # Дополнительная проверка даты выдачи паспорта
            if passport_issue_date > date.today():
                flash('Дата выдачи паспорта не может быть в будущем', 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)

            # Проверка возраста на дату выдачи паспорта
            age_at_issue = passport_issue_date.year - birth_date.year - (
                    (passport_issue_date.month, passport_issue_date.day) <
                    (birth_date.month, birth_date.day)
            )
            if age_at_issue < 14:
                flash('Паспорт не может быть выдан лицу младше 14 лет', 'error')
                return render_template('authors/create.html',
                                       today=date.today().isoformat(),
                                       education_levels=education_levels,
                                       form_data=request.form)

            # Создание автора
            author_id = author_dao.create(form_data)
            flash('Автор успешно создан!', 'success')
            return redirect(url_for('authors_list'))

        except ValueError as e:
            flash(f'Неверный формат даты: {str(e)}', 'error')
            return render_template('authors/create.html',
                                   today=date.today().isoformat(),
                                   education_levels=education_levels,
                                   form_data=request.form)
        except Exception as e:
            flash(f'Ошибка: {str(e)}', 'error')
            return render_template('authors/create.html',
                                   today=date.today().isoformat(),
                                   education_levels=education_levels,
                                   form_data=request.form)

    # GET запрос - пустая форма
    return render_template('authors/create.html',
                           today=date.today().isoformat(),
                           education_levels=education_levels,
                           form_data={})


@app.route('/author/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def author_edit(id):
    # Получаем автора
    author = author_dao.get_by_id(id)

    if not author:
        flash('Автор не найден', 'error')
        return redirect(url_for('authors_list'))

    # Проверяем права доступа
    author_user_id = author[12]  # user_id — 12‑й столбец (индекс 11)
    if session['role'] != 'admin' and author_user_id != session['user_id']:
        flash('У вас нет прав для редактирования этого автора', 'error')
        return redirect(url_for('authors_list'))

    if request.method == 'POST':
        try:
            # Подготавливаем даты
            birth_date = datetime.strptime(request.form['birth_date'], '%Y-%m-%d').date()

            # Подготавливаем дату выдачи паспорта (обязательное поле)
            passport_issue_date_str = request.form.get('passport_issue_date')
            if not passport_issue_date_str:
                flash('Дата выдачи паспорта обязательна для заполнения', 'error')
                return redirect(url_for('author_edit', id=id))

            passport_issue_date = datetime.strptime(passport_issue_date_str, '%Y-%m-%d').date()

            # Обработка чекбокса студента
            is_student = 'is_student' in request.form
            education_level_id = request.form.get('education_level_id')

            # Если студент, проверяем уровень образования
            if is_student and not education_level_id:
                flash('Для студента необходимо указать уровень образования', 'error')
                return redirect(url_for('author_edit', id=id))

            author_data = {
                'surname': request.form['surname'],
                'name': request.form['name'],
                'patronymic': request.form.get('patronymic', ''),
                'birth_date': birth_date,
                'address': request.form['address'],
                'department': request.form.get('department', ''),
                'passport': request.form['passport'],
                'passport_issue_date': passport_issue_date,
                'passport_issued_by': request.form.get('passport_issued_by', ''),
                'passport_issued_code': request.form.get('passport_issued_code', ''),
                'snils': request.form['snils'],
                'inn': request.form['inn'],
                'is_student': is_student,
                'position': request.form.get('position', ''),
                'citizenship': request.form.get('citizenship', 'Российская Федерация'),
                'education_level_id': int(education_level_id) if education_level_id else None
            }

            # Валидация через AuthorService
            errors = author_service.validate_author(author_data, id)

            # Дополнительная проверка даты выдачи паспорта
            if passport_issue_date > date.today():
                errors.append('Дата выдачи паспорта не может быть в будущем')

            # Проверка возраста на дату выдачи паспорта (временное решение)
            age_at_issue = passport_issue_date.year - birth_date.year - (
                    (passport_issue_date.month, passport_issue_date.day) <
                    (birth_date.month, birth_date.day)
            )
            if age_at_issue < 14:
                errors.append('Паспорт не может быть выдан лицу младше 14 лет')

            if errors:
                for error in errors:
                    flash(error, 'error')
                return redirect(url_for('author_edit', id=id))

            # Обновление автора
            success = author_dao.update(id, author_data)
            if success:
                flash('Автор успешно обновлен!', 'success')
                return redirect(url_for('authors_list'))
            else:
                flash('Не удалось обновить автора', 'error')
                return redirect(url_for('author_edit', id=id))

        except ValueError:
            flash('Неверный формат даты', 'error')
            return redirect(url_for('author_edit', id=id))
        except Exception as e:
            flash(f'Ошибка: {str(e)}', 'error')
            return redirect(url_for('author_edit', id=id))

    education_levels = author_dao.get_all_education_levels()

    today = date.today().isoformat()

    return render_template(
        'authors/edit.html',
        author=author,
        today=today,
        education_levels=education_levels or []
    )


@app.route('/author/<int:id>/delete', methods=['POST'])
@login_required
def author_delete(id):
    # Проверяем права доступа
    owner_id = author_dao.get_by_user(id, session['user_id'])
    if not owner_id:
        flash('Автор не найден', 'error')
        return redirect(url_for('authors_list'))

    if session['role'] != 'admin' and owner_id != session['user_id']:
        flash('У вас нет прав для удаления этого автора', 'error')
        return redirect(url_for('authors_list'))

    try:
        author_dao.delete(id)
        flash('Автор успешно удален', 'success')
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')

    return redirect(url_for('authors_list'))


@app.route('/author/<int:id>')
@login_required
def author_view(id):
    """Просмотр автора"""
    # Получаем автора
    author = author_dao.get_by_id(id)

    if not author:
        flash('Автор не найден', 'error')
        return redirect(url_for('authors_list'))

    # Преобразуем данные автора для шаблона
    author_list = list(author)

    # Преобразуем строки дат в объекты date
    for i in [4, 15]:  # birth_date, passport_issue_date
        if author_list[i] and isinstance(author_list[i], str):
            try:
                author_list[i] = datetime.strptime(author_list[i], '%Y-%m-%d').date()
            except ValueError:
                author_list[i] = None

    # Проверяем права доступа к автору
    author_user_id = author_list[12] if len(author_list) > 12 else author_list[12] if len(author_list) > 12 else None

    if not author_user_id:
        flash('Автор не найден', 'error')
        return redirect(url_for('authors_list'))

    # Пользователь может смотреть только своих авторов, админ - всех
    if session['role'] != 'admin' and author_user_id != session['user_id']:
        flash('У вас нет прав для просмотра этого автора', 'error')
        return redirect(url_for('authors_list'))

    return render_template('authors/view.html', author=author_list)


# Вспомогательная функция для получения уровней образования
def get_education_levels():
    """Получение всех уровней образования"""
    try:
        return author_dao.get_all_education_levels()
    except Exception as e:
        print(f"Error getting education levels: {e}")
        return []


# ========== ПРОДУКТЫ ==========
@app.route('/products')
@login_required
def products_list():
    if session['role'] == 'admin':
        products = product_dao.get_all()
    else:
        products = product_dao.get_all(session['user_id'])

    return render_template('products/list.html', products=products)


@app.route('/product/create', methods=['GET', 'POST'])
@login_required
def product_create():
    if request.method == 'POST':
        try:
            product_data = {
                'name': request.form['name'],
                'description': request.form['description'],
                'computer_type_id': request.form['computer_type_id'],
                'programming_language_id': request.form['programming_language_id'],
                'operating_system_id': request.form['operating_system_id'],
                'volume': request.form['volume'],
                'user_id': session['user_id']
            }

            product_id = product_dao.create(product_data)
            flash('Продукт успешно создан', 'success')
            return redirect(url_for('products_list'))

        except Exception as e:
            flash(f'Ошибка: {str(e)}', 'error')
            return render_template('products/create.html',
                                   computer_types=get_computer_types(),
                                   programming_languages=get_programming_languages(),
                                   os_types=get_os_types())

    return render_template('products/create.html',
                           computer_types=get_computer_types(),
                           programming_languages=get_programming_languages(),
                           os_types=get_os_types())


@app.route('/product/<int:id>')
@login_required
def product_view(id):
    # Проверяем права доступа
    owner_id = product_dao.get_by_user(id, session['user_id'])
    if not owner_id:
        flash('Продукт не найден', 'error')
        return redirect(url_for('products_list'))

    if session['role'] != 'admin' and owner_id != session['user_id']:
        flash('У вас нет прав для просмотра этого продукта', 'error')
        return redirect(url_for('products_list'))

    product = product_dao.get_by_id(id)
    if not product:
        flash('Продукт не найден', 'error')
        return redirect(url_for('products_list'))

    return render_template('products/view.html', product=product)


@app.route('/product/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def product_edit(id):
    # Проверяем права доступа
    owner_id = product_dao.get_by_user(id, session['user_id'])
    if not owner_id:
        flash('Продукт не найден', 'error')
        return redirect(url_for('products_list'))

    if session['role'] != 'admin' and owner_id != session['user_id']:
        flash('У вас нет прав для редактирования этого продукта', 'error')
        return redirect(url_for('products_list'))

    if request.method == 'POST':
        try:
            product_data = {
                'name': request.form['name'],
                'description': request.form['description'],
                'computer_type_id': request.form['computer_type_id'],
                'programming_language_id': request.form['programming_language_id'],
                'operating_system_id': request.form['operating_system_id'],
                'volume': request.form['volume']
            }

            product_dao.update(id, product_data)
            flash('Продукт успешно обновлен', 'success')
            return redirect(url_for('products_list'))

        except Exception as e:
            flash(f'Ошибка: {str(e)}', 'error')

    product = product_dao.get_by_id(id)
    if not product:
        flash('Продукт не найден', 'error')
        return redirect(url_for('products_list'))

    return render_template('products/edit.html',
                           product=product,
                           computer_types=get_computer_types(),
                           programming_languages=get_programming_languages(),
                           os_types=get_os_types())


@app.route('/product/<int:id>/delete', methods=['POST'])
@login_required
def product_delete(id):
    # Проверяем права доступа
    owner_id = product_dao.get_by_user(id, session['user_id'])
    if not owner_id:
        flash('Продукт не найден', 'error')
        return redirect(url_for('products_list'))

    if session['role'] != 'admin' and owner_id != session['user_id']:
        flash('У вас нет прав для удаления этого продукта', 'error')
        return redirect(url_for('products_list'))

    try:
        product_dao.delete(id)
        flash('Продукт успешно удален', 'success')
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')

    return redirect(url_for('products_list'))


# ========== ЗАЯВКИ ==========
@app.route('/applications')
@login_required
def applications_list():
    application_dao = ApplicationDAO(db)  # Инициализируем DAO

    if session['role'] == 'admin':
        applications = application_dao.get_all()
    else:
        applications = application_dao.get_all(session['user_id'])

    all_statuses = get_application_statuses()

    return render_template('applications/list.html',
                           applications=applications,
                           all_statuses=all_statuses,
                           application_dao=application_dao)  # Передаем DAO в шаблон


@app.route('/application/create', methods=['GET', 'POST'])
@login_required
def application_create():
    application_dao = ApplicationDAO(db)

    if request.method == 'POST':
        try:
            # Проверяем сумму процентов участия
            percents = request.form.getlist('percents[]')
            total_percent = sum(int(percent) for percent in percents if percent)

            if total_percent != 100:
                flash(f'Сумма процентов участия должна быть 100%. Текущая сумма: {total_percent}%', 'error')
                return redirect(url_for('application_create'))

            # Создаем заявку с привязкой к пользователю и комиссии
            application_data = {
                'title': request.form['title'],
                'product_id': request.form['product_id'],
                'user_id': session['user_id'],
                'commission_id': request.form.get('commission_id') or None,
                'nir': request.form['nir']
            }

            application_id = application_dao.create(application_data)

            # Добавляем авторов (только тех, которые принадлежат пользователю)
            author_ids = request.form.getlist('author_ids[]')
            role_ids = request.form.getlist('role_ids[]')
            percents = request.form.getlist('percents[]')
            times = request.form.getlist('times[]')
            descriptions = request.form.getlist('description[]')
            equipment = request.form.getlist('equipment[]')

            for i in range(len(author_ids)):
                # Проверяем, принадлежит ли автор пользователю
                owner_id = author_dao.get_by_user(int(author_ids[i]), session['user_id'])
                if owner_id or session['role'] == 'admin':
                    author_data = {
                        'author_id': author_ids[i],
                        'role_id': role_ids[i],
                        'participation_percent': percents[i],
                        'time_spent': times[i] if i < len(times) else None,
                        'description': descriptions[i] if i < len(descriptions) else '',
                        'equipment': equipment[i] if i < len(equipment) else '',
                    }
                    application_dao.add_author_to_application(application_id, author_data)

            flash('Заявка успешно создана', 'success')
            return redirect(url_for('applications_list'))

        except Exception as e:
            flash(f'Ошибка при создании заявки: {str(e)}', 'error')

    # Получаем данные для формы
    if session['role'] == 'admin':
        authors = author_dao.get_all()
        products = product_dao.get_all()
        commissions = commission_dao.get_all()
    else:
        authors = author_dao.get_all(session['user_id'])
        products = product_dao.get_all(session['user_id'])
        commissions = commission_dao.get_all(session['user_id'])

    roles = get_expert_roles()
    statuses = get_application_statuses()

    return render_template('applications/create.html',
                           authors=authors,
                           products=products,
                           commissions=commissions,
                           roles=roles,
                           statuses=statuses)


@app.route('/application/<int:id>')
@login_required
def application_view(id):
    """Просмотр одной заявки"""
    application_dao = ApplicationDAO(db)

    # Получаем заявку
    if session['role'] == 'admin':
        application = application_dao.get_by_id(id)
    else:
        application = application_dao.get_by_id(id, session['user_id'])

    if not application:
        flash('Заявка не найдена', 'error')
        return redirect(url_for('applications_list'))

    # Проверяем права доступа
    if session['role'] != 'admin':
        # У пользователя должно быть поле user_id в записи
        if len(application) > 4 and application[4] != session['user_id']:
            flash('У вас нет прав для просмотра этой заявки', 'error')
            return redirect(url_for('applications_list'))

    # Получаем авторов заявки
    application_authors = application_dao.get_application_authors(id)
    all_statuses = get_application_statuses()

    # Рассчитываем общее время и проценты
    total_time = 0
    total_percent = 0

    for author in application_authors:
        time_spent = author[5] if len(author) > 5 else 0
        participation_percent = author[4] if len(author) > 4 else 0

        if time_spent:
            total_time += time_spent
        if participation_percent:
            total_percent += participation_percent

    # Преобразуем дату создания
    created_date = None
    if application[3]:  # created_date
        if isinstance(application[3], (int, float)):
            created_date = datetime.fromtimestamp(application[3])
        elif isinstance(application[3], str):
            try:
                created_date = datetime.fromisoformat(application[3])
            except ValueError:
                created_date = application[3]
        elif hasattr(application[3], 'strftime'):
            created_date = application[3]

    if session['role'] == 'admin':
        product_name = application[8] if len(application) > 8 else 'Не указан'
        commission_name = application[10] if len(application) > 10 else ''
        status_name = application[11] if len(application) > 11 else 'Неизвестно'
        status_id = application[12] if len(application) > 12 else None
        username = application[9] if len(application) > 9 else ''
    else:
        product_name = application[8] if len(application) > 8 else 'Не указан'
        commission_name = application[9] if len(application) > 9 else ''
        status_name = application[10] if len(application) > 10 else 'Неизвестно'
        status_id = application[11] if len(application) > 11 else None
        username = application[12] if len(application) > 12 else ''

    return render_template('applications/view.html',
                           application=application,
                           product_name=product_name,
                           commission_name=commission_name,
                           status_name=status_name,
                           status_id=status_id,
                           username=username,
                           application_authors=application_authors,
                           total_time=total_time,
                           total_percent=total_percent,
                           created_date=created_date,
                           all_statuses=all_statuses,
                           application_dao=application_dao)


@app.route('/application/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def application_edit(id):
    """Редактирование заявки"""
    application_dao = ApplicationDAO(db)

    # Получаем заявку
    if session['role'] == 'admin':
        application = application_dao.get_by_id(id)
    else:
        application = application_dao.get_by_id(id, session['user_id'])

    if not application:
        flash('Заявка не найдена', 'error')
        return redirect(url_for('applications_list'))

    # Проверяем возможность редактирования
    if not application_dao.can_user_edit(id, session['user_id'], session['role'] == 'admin'):
        flash('Редактирование заявки недоступно в текущем статусе', 'error')
        return redirect(url_for('application_view', id=id))

    if request.method == 'POST':
        try:
            # Проверяем сумму процентов участия
            percents = request.form.getlist('percents[]')
            total_percent = sum(float(percent) for percent in percents if percent)
            if total_percent != 100:
                flash(f'Сумма процентов участия должна быть 100%. Текущая сумма: {total_percent}%', 'error')
                return redirect(url_for('application_edit', id=id))

            # Обновляем данные заявки
            update_data = {
                'title': request.form['title'],
                'product_id': request.form['product_id'],
                'commission_id': request.form.get('commission_id') or None,
                'nir': request.form['nir']
            }

            if application_dao.update_application(
                    id,
                    session['user_id'],
                    update_data,
                    session['role'] == 'admin'
            ):
                flash('Заявка успешно обновлена!', 'success')
            else:
                flash('Не удалось обновить заявку', 'error')

            # Обновляем существующих авторов
            author_ids = request.form.getlist('author_ids[]')
            role_ids = request.form.getlist('role_ids[]')
            percents = request.form.getlist('percents[]')
            times = request.form.getlist('times[]')
            descriptions = request.form.getlist('description[]')
            equipment = request.form.getlist('equipment[]')

            # Получаем текущих авторов заявки
            current_authors = application_dao.get_application_authors(id)
            current_author_ids = {str(author[0]) for author in current_authors}

            # Обновляем только существующих авторов
            for i in range(len(author_ids)):
                if author_ids[i] in current_author_ids:
                    author_data = {
                        'author_id': author_ids[i],
                        'role_id': role_ids[i],
                        'participation_percent': percents[i],
                        'time_spent': times[i] if i < len(times) else None,
                        'description': descriptions[i] if i < len(descriptions) else '',
                        'equipment': equipment[i] if i < len(equipment) else '',
                    }
                    application_dao.update_application_author(id, author_ids[i], author_data)

            return redirect(url_for('application_view', id=id))

        except Exception as e:
            flash(f'Ошибка при обновлении заявки: {str(e)}', 'error')

    # Получаем текущих авторов заявки
    application_authors = application_dao.get_application_authors(id)

    # Получаем данные для формы
    if session['role'] == 'admin':
        # Берем всех авторов, чтобы отобразить их в форме
        authors = author_dao.get_all()
        products = product_dao.get_all()
        commissions = commission_dao.get_all()
    else:
        authors = author_dao.get_all(session['user_id'])
        products = product_dao.get_all(session['user_id'])
        commissions = commission_dao.get_all(session['user_id'])

    roles = get_expert_roles()
    statuses = get_application_statuses()

    return render_template('applications/edit.html',
                           application=application,
                           application_authors=application_authors,
                           authors=authors,
                           products=products,
                           commissions=commissions,
                           roles=roles,
                           statuses=statuses)


@app.route('/application/<int:id>/submit_for_review', methods=['POST'])
@login_required
def application_submit_for_review(id):
    """Отправка заявки на рассмотрение"""
    application_dao = ApplicationDAO(db)

    # Проверяем права доступа
    if session['role'] != 'admin':
        app_user_id = application_dao.get_by_user(id, session['user_id'])
        if not app_user_id or app_user_id != session['user_id']:
            flash('У вас нет прав для выполнения этого действия', 'error')
            return redirect(url_for('application_view', id=id))

    # Отправляем на рассмотрение
    if application_dao.submit_for_review(id, session['user_id']):
        flash('Заявка успешно отправлена на рассмотрение', 'success')
    else:
        flash('Не удалось отправить заявку на рассмотрение. Проверьте статус заявки', 'error')

    return redirect(url_for('application_view', id=id))


@app.route('/application/<int:id>/delete', methods=['POST'])
@login_required
def application_delete(id):
    application_dao = ApplicationDAO(db)

    # Проверяем права доступа
    owner_id = application_dao.get_by_user(id, session['user_id'])
    if not owner_id:
        flash('Заявка не найдена', 'error')
        return redirect(url_for('applications_list'))

    if session['role'] != 'admin' and owner_id != session['user_id']:
        flash('У вас нет прав для удаления этой заявки', 'error')
        return redirect(url_for('applications_list'))

    try:
        application_dao.delete(id)
        flash('Заявка успешно удалена', 'success')
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')

    return redirect(url_for('applications_list'))


@app.route('/application/<int:id>/update_status', methods=['POST'])
@login_required
def application_update_status(id):
    application_dao = ApplicationDAO(db)

    # Только админ может менять статусы
    if session['role'] != 'admin':
        flash('Только администратор может менять статусы заявок', 'error')
        return redirect(url_for('application_view', id=id))

    new_status_id = request.form['status_id']

    # Проверяем, что статус существует
    status = get_application_status_by_id(new_status_id)
    if not status:
        flash('Неверный статус', 'error')
        return redirect(url_for('application_view', id=id))

    # Обновляем статус
    try:
        application_dao.update_status(id, new_status_id)
        flash(f'Статус заявки изменен на "{status["name"]}"', 'success')
    except Exception as e:
        flash(f'Ошибка при обновлении статуса: {str(e)}', 'error')

    return redirect(url_for('application_view', id=id))


# ========== ЭКСПЕРТНЫЕ КОМИССИИ ==========
@app.route('/commissions')
@login_required
def commissions_list():
    """Список экспертных комиссий"""
    if session['role'] == 'admin':
        commissions = commission_dao.get_all()
    else:
        commissions = commission_dao.get_all(session['user_id'])

    return render_template('commissions/list.html', commissions=commissions)


@app.route('/commission/<int:id>', methods=['GET'])
@login_required
def commission_view(id):
    """Просмотр комиссии"""
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    # Проверяем доступ
    # commission[4] - это user_id (4-й элемент в кортеже)
    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет доступа к этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    # Получаем экспертов комиссии
    commission_experts = commission_dao.get_commission_experts(id)

    # Получаем доступных экспертов для добавления (только эксперты текущего пользователя)
    available_experts = commission_dao.get_available_experts_for_commission(id, session['user_id'])

    # Получаем роли
    roles = commission_dao.get_commission_roles()

    # Преобразуем комиссию в словарь для удобства
    commission_dict = {
        'id': commission[0],
        'name': commission[1],
        'description': commission[2],
        'created_date': commission[3],
        'user_id': commission[4],
        'username': commission[5] if len(commission) > 5 else None,
        'member_count': commission[6] if len(commission) > 6 else 0
    }

    return render_template('commissions/view.html',
                           commission=commission_dict,
                           commission_experts=commission_experts,
                           available_experts=available_experts,
                           roles=roles)


@app.route('/expert/create', methods=['POST'])
@login_required
def expert_create():
    """Создание нового эксперта"""
    surname = request.form.get('surname')
    name = request.form.get('name')
    patronymic = request.form.get('patronymic')
    position = request.form.get('position')

    if not all([surname, name, position]):
        flash('Заполните обязательные поля', 'error')
        return redirect(url_for('experts_list'))

    expert_data = {
        'surname': surname,
        'name': name,
        'patronymic': patronymic,
        'position': position,
        'user_id': session['user_id']
    }

    try:
        expert_dao.create(expert_data)
        flash('Эксперт успешно создан', 'success')
    except Exception as e:
        flash(f'Ошибка при создании эксперта: {str(e)}', 'error')

    return redirect(url_for('experts_list'))


@app.route('/commission/<int:id>/add_expert', methods=['POST'])
@login_required
def commission_add_expert(id):
    """Добавление эксперта в комиссию"""
    expert_id = request.form.get('expert_id')
    role_id = request.form.get('role_id')

    if not all([expert_id, role_id]):
        flash('Выберите эксперта и роль', 'error')
        return redirect(url_for('commission_view', id=id))

    # Проверяем доступ к комиссии
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет доступа к этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    # Проверяем доступ к эксперту
    expert = expert_dao.get_by_id(int(expert_id))
    if not expert:
        flash('Эксперт не найден', 'error')
        return redirect(url_for('commission_view', id=id))

    if session['role'] != 'admin' and expert[5] != session['user_id']:
        flash('У вас нет доступа к этому эксперту', 'error')
        return redirect(url_for('commission_view', id=id))

    # Добавляем эксперта в комиссию
    if commission_dao.add_expert_to_commission(id, int(expert_id), int(role_id)):
        flash('Эксперт успешно добавлен в комиссию', 'success')
    else:
        flash('Эксперт уже состоит в этой комиссии', 'error')

    return redirect(url_for('commission_view', id=id))


@app.route('/commission/<int:id>/remove_expert/<int:expert_id>', methods=['POST'])
@login_required
def commission_remove_expert(id, expert_id):
    """Удаление эксперта из комиссии"""
    # Проверяем доступ к комиссии
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет доступа к этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    # Удаляем эксперта из комиссии
    if commission_dao.remove_expert_from_commission(id, expert_id):
        flash('Эксперт успешно удален из комиссии', 'success')
    else:
        flash('Не удалось удалить эксперта из комиссии', 'error')

    return redirect(url_for('commission_view', id=id))


@app.route('/commission/<int:id>/delete', methods=['POST'])
@login_required
def commission_delete(id):
    """Удаление экспертной комиссии"""
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    # Проверяем права доступа
    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет прав для удаления этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    try:
        success = commission_dao.delete(id)
        if success:
            flash('Экспертная комиссия успешно удалена!', 'success')
        else:
            flash('Не удалось удалить комиссию', 'error')
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')

    return redirect(url_for('commissions_list'))


@app.route('/commission/<int:id>/update_expert_role', methods=['POST'])
@login_required
def commission_update_expert_role(id):
    """Обновление роли эксперта в комиссии"""
    expert_id = request.form.get('expert_id')
    role_id = request.form.get('role_id')

    if not all([expert_id, role_id]):
        flash('Не указаны эксперт или роль', 'error')
        return redirect(url_for('commission_view', id=id))

    # Проверяем доступ к комиссии
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет доступа к этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    # Обновляем роль
    if commission_dao.update_expert_role(id, int(expert_id), int(role_id)):
        flash('Роль эксперта успешно обновлена', 'success')
    else:
        flash('Не удалось обновить роль эксперта', 'error')

    return redirect(url_for('commission_view', id=id))


@app.route('/experts')
@login_required
def experts_list():
    """Список экспертов"""
    if session['role'] == 'admin':
        experts = expert_dao.get_all()
    else:
        experts = expert_dao.get_all(session['user_id'])

    return render_template('experts/list.html', experts=experts)


@app.route('/expert/<int:id>/delete', methods=['POST'])
@login_required
def expert_delete(id):
    """Удаление эксперта"""
    expert = expert_dao.get_by_id(id)
    if not expert:
        flash('Эксперт не найден', 'error')
        return redirect(url_for('experts_list'))

    # Проверяем права доступа
    if session['role'] != 'admin' and expert[5] != session['user_id']:
        flash('У вас нет прав для удаления этого эксперта', 'error')
        return redirect(url_for('experts_list'))

    try:
        success = expert_dao.delete(id)
        if success:
            flash('Эксперт успешно удален!', 'success')
        else:
            flash('Не удалось удалить эксперта (возможно, он состоит в комиссиях)', 'error')
    except Exception as e:
        flash(f'Ошибка: {str(e)}', 'error')

    return redirect(url_for('experts_list'))


@app.route('/commissions/create', methods=['GET', 'POST'])
@login_required
def commission_create():
    """Создание новой комиссии с экспертами"""
    if request.method == 'POST':
        name = request.form.get('name')
        description = request.form.get('description', '')

        # Получаем данные о выбранных существующих экспертах
        selected_experts_data = request.form.get('selected_experts_data')
        # Получаем данные о новых экспертах
        new_experts_data = request.form.get('new_experts_data')

        if not name:
            flash('Введите название комиссии', 'error')
            return redirect(url_for('commission_create'))

        try:
            # Создаем комиссию
            commission_id = commission_dao.create(name, description, session['user_id'])

            # Обрабатываем существующих экспертов
            if selected_experts_data:
                import json
                experts = json.loads(selected_experts_data)
                for expert in experts:
                    commission_dao.add_expert_to_commission(
                        commission_id,
                        int(expert['expert_id']),
                        int(expert['role_id'])
                    )

            # Обрабатываем новых экспертов
            if new_experts_data:
                import json
                experts = json.loads(new_experts_data)
                for expert_data in experts:
                    # Создаем нового эксперта
                    expert_dict = {
                        'surname': expert_data['surname'],
                        'name': expert_data['name'],
                        'patronymic': expert_data.get('patronymic', ''),
                        'position': expert_data['position'],
                        'user_id': session['user_id']
                    }

                    # Создаем эксперта и получаем его ID
                    expert_id = expert_dao.create(expert_dict)

                    # Добавляем в комиссию
                    commission_dao.add_expert_to_commission(
                        commission_id,
                        expert_id,
                        int(expert_data['role_id'])
                    )

            flash('Комиссия успешно создана!', 'success')
            return redirect(url_for('commissions_list', id=commission_id))
        except Exception as e:
            flash(f'Ошибка при создании комиссии: {str(e)}', 'error')
            # Получаем доступных экспертов для повторного отображения формы
            available_experts = commission_dao.get_available_experts_for_commission(None, session['user_id'])
            roles = commission_dao.get_commission_roles()
            return render_template('commissions/create.html',
                                   available_experts=available_experts,
                                   commission_roles=roles)

    # GET запрос - показать форму
    # Получаем доступных экспертов текущего пользователя
    available_experts = commission_dao.get_available_experts_for_commission(None, session['user_id'])
    roles = commission_dao.get_commission_roles()

    return render_template('commissions/create.html',
                           available_experts=available_experts,
                           commission_roles=roles)


@app.route('/experts/create', methods=['GET', 'POST'])
@login_required
def expert_create_page():
    """Страница создания нового эксперта"""
    if request.method == 'POST':
        surname = request.form.get('surname')
        name = request.form.get('name')
        patronymic = request.form.get('patronymic', '')
        position = request.form.get('position')

        if not all([surname, name, position]):
            flash('Заполните обязательные поля: фамилия, имя, должность', 'error')
            return render_template('experts/create.html')

        expert_data = {
            'surname': surname,
            'name': name,
            'patronymic': patronymic,
            'position': position,
            'user_id': session['user_id']
        }

        try:
            expert_id = expert_dao.create(expert_data)
            flash('Эксперт успешно создан!', 'success')
            return redirect(url_for('expert_view', id=expert_id))
        except Exception as e:
            flash(f'Ошибка при создании эксперта: {str(e)}', 'error')
            return render_template('experts/create.html')

    # GET запрос - показать форму создания
    return render_template('experts/create.html')


# ========== ГЕНЕРАЦИЯ ДОКУМЕНТОВ ==========
@app.route('/application/<int:id>/generate_consent')
def generate_consent(id):
    """Генерация документов согласия для всех авторов - возвращает ZIP-архив со всеми документами"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        print(f"DEBUG: Starting all-authors consent generation for application {id}")

        # Получаем заявку
        application = application_dao.get_by_id(id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=id))

        # Получаем авторов
        authors_data = application_dao.get_application_authors_detailed(id)
        if not authors_data:
            flash('В заявке нет авторов', 'error')
            return redirect(url_for('application_view', id=id))

        # Получаем настройки
        settings_data = setting_dao.get_current()

        # Создаем генератор
        generator = ConsentGenerator()

        # Получаем ВСЕ документы
        all_documents = generator.generate_all_authors(application, authors_data, settings_data)

        if not all_documents:
            flash('Не удалось сгенерировать документы', 'error')
            return redirect(url_for('application_view', id=id))

        # Создаем ZIP-архив в памяти
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, doc_buffer in enumerate(all_documents):
                # Получаем автора для формирования имени файла
                if i < len(authors_data):
                    author = authors_data[i]
                    if len(author) > 7:
                        surname = str(author[7]).strip() if author[7] else ''
                        name = str(author[8]).strip() if author[8] else ''
                        patronymic = str(author[9]).strip() if author[9] else ''

                        # Очищаем имя файла от недопустимых символов
                        def clean_filename(text):
                            if not text:
                                return ""
                            # Убираем недопустимые символы для Windows/Linux
                            invalid_chars = '<>:"/\\|?*'
                            for char in invalid_chars:
                                text = text.replace(char, '')
                            return text.strip()

                        surname_clean = clean_filename(surname)
                        name_clean = clean_filename(name)

                        if surname_clean and name_clean:
                            if patronymic:
                                filename = f"sogl_na_ykaz_dan_zayavka_author_{surname_clean}_{name_clean}.docx"
                            else:
                                filename = f"sogl_na_ykaz_dan_zayavka_author_{surname_clean}_{name_clean}.docx"
                        else:
                            filename = f"sogl_na_ykaz_dan_zayavka_author_{i + 1}.docx"
                    else:
                        filename = f"sogl_na_ykaz_dan_zayavka_author_{i + 1}.docx"
                else:
                    filename = f"sogl_na_ykaz_dan_zayavka_author_{i + 1}.docx"

                # Добавляем документ в ZIP-архив
                zip_file.writestr(filename, doc_buffer.getvalue())

        zip_buffer.seek(0)

        # Имя для ZIP-архива
        zip_filename = f"sogl_na_ykaz_dan_zayavka_{id}.zip"

        # Отправляем ZIP-архив пользователю
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации ZIP с согласиями: {e}", exc_info=True)
        flash(f'Ошибка генерации документов: {str(e)}', 'error')
        return redirect(url_for('application_view', id=id))


@app.route('/application/<int:application_id>/generate_abstract')
def generate_abstract_from_application(application_id):
    """Генерация аннотации из заявки"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем ID продукта из заявки
        product_id = application[2]  # product_id находится в индексе 2

        # Получаем данные продукта
        product_data = product_dao.get_by_id(product_id)
        if not product_data:
            flash('Продукт не найден', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов с ПОДРОБНЫМИ данными (get_application_authors_detailed)
        authors_data = application_dao.get_application_authors_detailed(application_id)

        print(f"DEBUG: Authors data type: {type(authors_data)}")
        print(f"DEBUG: Authors data length: {len(authors_data) if authors_data else 0}")
        if authors_data and len(authors_data) > 0:
            print(f"DEBUG: First author: {authors_data[0]}")
            print(f"DEBUG: First author length: {len(authors_data[0])}")

        # Создаем генератор
        generator = AbstractGenerator()
        doc = generator.generate(product_data, authors_data)

        if doc:
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            filename = f"referat_zayavka_{application_id}.docx"

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('Не удалось сгенерировать документ', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации реферата: {e}", exc_info=True)
        flash(f'Ошибка генерации документа: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_sogl')
def generate_sogl_from_application(application_id):
    """Генерация отдельных документов для каждого автора"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов заявки
        application_authors = application_dao.get_application_authors(application_id)

        if not application_authors:
            flash('Не найдены авторы заявки', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Создаем генератор
        generator = SoglGenerator()

        # Генерируем документы
        documents = generator.generate_for_each_author(application, application_authors, author_dao)

        if not documents:
            flash('Ошибка генерации согласия на обработку персональных данных', 'error')
            return redirect(url_for('application_view', id=application_id))

        if len(documents) == 1:
            # Один автор - один файл
            file_stream = io.BytesIO()
            documents[0].save(file_stream)
            file_stream.seek(0)

            # Имя файла с ASCII символами
            filename = f"sogl_{application_id}.docx"

            response = Response(
                file_stream.read(),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={
                    'Content-Disposition': f'attachment; filename="{filename}"'
                }
            )
            return response

        else:
            # Несколько авторов - ZIP архив
            import zipfile

            zip_stream = io.BytesIO()
            with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i, doc in enumerate(documents):
                    # Сохраняем документ
                    doc_stream = io.BytesIO()
                    doc.save(doc_stream)
                    doc_stream.seek(0)

                    # Имя файла с ASCII символами
                    filename = f"sogl_{application_id}_author_{i + 1}.docx"
                    zip_file.writestr(filename, doc_stream.read())

            zip_stream.seek(0)

            filename = f"sogl_{application_id}.zip"

            response = Response(
                zip_stream.read(),
                mimetype='application/zip',
                headers={
                    'Content-Disposition': f'attachment; filename="{filename}"'
                }
            )
            return response

    except Exception as e:
        print(f"Ошибка при генерации согласия: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при генерации согласия: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_conclusion')
def generate_conclusion_from_application(application_id):
    """Генерация заключения о возможности открытого опубликования из заявки"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        user_id = session['user_id']
        print(f"DEBUG: User ID from session: {user_id}")

        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('applications_list'))

        print(f"DEBUG: Application tuple: {application}")
        print(f"DEBUG: Application length: {len(application)}")

        # ВРЕМЕННО ОТКЛЮЧАЕМ ПРОВЕРКУ ДОСТУПА ДЛЯ ОТЛАДКИ
        access_granted = True
        print("DEBUG: Access check disabled for debugging")

        if not access_granted:
            flash('Нет доступа к этой заявке', 'error')
            return redirect(url_for('applications_list'))

        # ПОЛУЧАЕМ COMMISSION_ID - ВАРИАНТ 1: Из заявки
        commission_id = None
        print(f"DEBUG: Checking all application fields for commission_id:")

        # Проверяем разные возможные индексы
        indices_to_check = [12, 13, 14, 15, 16, 17]
        for idx in indices_to_check:
            if len(application) > idx:
                field = application[idx]
                print(f"  Index {idx}: {field} (type: {type(field).__name__})")
                if field and isinstance(field, (int, float, str)):
                    if str(field).isdigit():
                        commission_id = int(field)
                        print(f"  ✓ Found possible commission_id at index {idx}: {commission_id}")
                        # Не прерываем, чтобы увидеть все возможные места

        # ВАРИАНТ 2: Проверить прямо в БД
        print(f"DEBUG: Checking commission_id directly in database...")
        with db.get_cursor() as cur:
            # Сначала проверим структуру таблицы
            cur.execute("""
                SELECT column_name 
                FROM information_schema.columns 
                WHERE table_name = 'application' 
                ORDER BY ordinal_position
            """)
            columns = [col[0] for col in cur.fetchall()]
            print(f"DEBUG: Application table columns: {columns}")

            # Ищем столбец с commission
            for col in columns:
                if 'commission' in col.lower():
                    cur.execute(f"SELECT {col} FROM application WHERE id = %s", (application_id,))
                    result = cur.fetchone()
                    if result:
                        db_commission_id = result[0]
                        print(f"DEBUG: Found commission_id in column '{col}': {db_commission_id}")
                        if db_commission_id and commission_id != db_commission_id:
                            commission_id = db_commission_id
                            print(f"  ✓ Using commission_id from DB column '{col}': {commission_id}")

        if not commission_id:
            print(f"DEBUG: No commission_id found, trying default commission...")
            # Можно попробовать получить первую комиссию пользователя
            user_commissions = commission_dao.get_all(user_id)
            if user_commissions and len(user_commissions) > 0:
                commission_id = user_commissions[0][0]
                print(f"DEBUG: Using first user commission: {commission_id}")

        print(f"DEBUG: Final commission_id to use: {commission_id}")

        commission_experts = None

        if commission_id:
            # Получаем экспертов комиссии
            print(f"DEBUG: Getting experts for commission_id: {commission_id}")
            commission_experts = commission_dao.get_commission_experts(commission_id)
            print(f"DEBUG: Commission experts count: {len(commission_experts) if commission_experts else 0}")

            if commission_experts and len(commission_experts) > 0:
                print(f"DEBUG: Commission experts data sample: {commission_experts[0]}")
                print(f"DEBUG: Expert tuple length: {len(commission_experts[0])}")

                # Проверяем, что это полные данные (должно быть минимум 6 полей)
                if len(commission_experts[0]) >= 6:
                    print(f"DEBUG: Commission experts have full data structure")
                    # Выводим структуру первого эксперта
                    for i, field in enumerate(commission_experts[0]):
                        print(f"DEBUG: Expert[0][{i}]: {field}")
                else:
                    print(f"DEBUG: WARNING: Commission experts have insufficient data")
                    print(f"DEBUG: Trying to get detailed data...")

                    # Пробуем получить подробные данные
                    with db.get_cursor() as cur:
                        cur.execute("""
                            SELECT 
                                e.id, 
                                e.surname, 
                                e.name, 
                                e.patronymic, 
                                e.position,
                                cr.name as role_name
                            FROM commission_experts ce
                            JOIN expert e ON ce.expert_id = e.id
                            JOIN commission_roles cr ON ce.role_id = cr.id
                            WHERE ce.commission_id = %s
                            ORDER BY 
                                CASE cr.name
                                    WHEN 'председатель' THEN 1
                                    WHEN 'эксперт' THEN 2
                                    WHEN 'секретарь' THEN 3
                                    ELSE 4
                                END,
                                e.surname, e.name, e.patronymic
                        """, (commission_id,))
                        detailed_experts = cur.fetchall()
                        if detailed_experts and len(detailed_experts) > 0:
                            commission_experts = detailed_experts
                            print(f"DEBUG: Got detailed experts: {len(commission_experts)} records")
                            print(f"DEBUG: First detailed expert: {commission_experts[0]}")
            else:
                print(f"DEBUG: No experts found in commission {commission_id}")

                # Покажем, какие эксперты вообще есть в системе
                with db.get_cursor() as cur:
                    cur.execute("SELECT COUNT(*) FROM expert WHERE user_id = %s", (user_id,))
                    expert_count = cur.fetchone()[0]
                    print(f"DEBUG: User has {expert_count} experts in system")

                    cur.execute("SELECT id, surname, name, patronymic FROM expert WHERE user_id = %s LIMIT 5",
                                (user_id,))
                    sample_experts = cur.fetchall()
                    print(f"DEBUG: Sample experts: {sample_experts}")
        else:
            print("DEBUG: No commission ID determined")

        # Получаем текущие настройки
        settings_data = setting_dao.get_current()
        print(f"DEBUG: Settings data: {settings_data}")

        # Получаем авторов заявки
        authors_data = application_dao.get_application_authors_detailed(application_id)
        print(f"DEBUG: Authors data count: {len(authors_data) if authors_data else 0}")

        # Создаем генератор заключения
        generator = ConclusionGenerator()

        # Если нет данных комиссии, создаем тестовые данные
        if not commission_experts:
            print("DEBUG: No commission experts, using TEST data")
            # Тестовые данные для демонстрации
            test_experts = [
                (1, 'Иванов', 'Иван', 'Иванович', 'профессор кафедры информатики', 'председатель'),
                (2, 'Петрова', 'Мария', 'Сергеевна', 'доцент', 'эксперт'),
                (3, 'Сидоров', 'Алексей', 'Петрович', 'старший преподаватель', 'эксперт')
            ]
            commission_experts = test_experts
            print(f"DEBUG: Using test experts data")

        # Генерируем документ с передачей данных авторов
        doc = generator.generate(
            application=application,
            commission_data=commission_experts,
            settings_data=settings_data,
            authors_data=authors_data
        )

        if doc:
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            # Получаем название продукта для имени файла
            product_name = generator._extract_product_name(application)
            safe_product_name = "".join(c for c in product_name if c.isalnum() or c in (' ', '-', '_')).rstrip()

            if not safe_product_name:
                safe_product_name = f"zayavka_{application_id}"

            # Ограничиваем длину имени файла
            if len(safe_product_name) > 50:
                safe_product_name = safe_product_name[:50]

            filename = f"zakl_o_vozm_otkr_opybl_{safe_product_name}.docx"

            print(f"DEBUG: Generated filename: {filename}")

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('Не удалось сгенерировать документ заключения о возможности открытого опубликования', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации заключения: {e}", exc_info=True)
        flash(f'Ошибка генерации документа: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_distribution_agreement')
def generate_distribution_agreement(application_id):
    """Генерация соглашения о распределении вознаграждения между авторами"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов заявки
        authors_data = application_dao.get_application_authors_detailed(application_id)
        if not authors_data:
            flash('У заявки нет авторов', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Создаем генератор соглашения
        generator = DistributionAgreementGenerator(template_name='distributionagreement_generator.docx')

        # Генерируем документ
        doc = generator.generate(application, authors_data)

        if doc:
            # Сохраняем в поток
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            # Простое ASCII имя файла
            filename = f"sogl_o_raspr_vozn_{application_id}.docx"

            # Используем Response с ASCII именем файла
            response = Response(
                file_stream.read(),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={
                    'Content-Disposition': f'attachment; filename="{filename}"'
                }
            )

            return response

        else:
            flash('Не удалось сгенерировать документ', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        # Более безопасный вывод ошибки
        error_msg = str(e)
        print(f"ERROR in generate_distribution_agreement: {error_msg}")
        import traceback
        traceback.print_exc()

        flash('Произошла ошибка при генерации документа. Обратитесь к администратору.', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_rid_notification')
def generate_rid_notification(application_id):
    """Генерация уведомления о создании РИД (простая версия)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов заявки
        authors_data = application_dao.get_application_authors_detailed(application_id)
        if not authors_data:
            flash('У заявки нет авторов', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Создаем генератор уведомления
        generator = RIDNotificationGenerator(template_name='notification_template.docx')

        if hasattr(generator, 'set_daos'):
            generator.set_daos(setting_dao, application_dao)

        # Генерируем документ
        doc = generator.generate(application_id, session['user_id'])

        if doc:
            # Сохраняем в поток
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            # ПРОСТОЕ ASCII ИМЯ ФАЙЛА - чтобы избежать проблем с кодировкой
            filename = f"yved_rid_{application_id}.docx"

            # Используем Response с ASCII именем файла
            response = Response(
                file_stream.read(),
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={
                    'Content-Disposition': f'attachment; filename="{filename}"'
                }
            )

            return response

        else:
            flash('Не удалось сгенерировать уведомление', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        error_msg = str(e)
        print(f"ERROR in generate_rid_notification: {error_msg}")
        import traceback
        traceback.print_exc()

        flash('Произошла ошибка при генерации уведомления. Обратитесь к администратору.', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_employee_assignment_agreements')
def generate_assignment_agreement(application_id):
    """Генерация договоров об отчуждении исключительного права только для сотрудников"""
    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(f'/application/{application_id}')

        # Получаем авторов заявки
        authors_data = application_dao.get_application_authors_detailed(application_id)
        if not authors_data:
            flash('У заявки нет авторов', 'error')
            return redirect(f'/application/{application_id}')

        # Создаем генератор
        generator = AssignmentAgreementGenerator(template_name='assignment_agreement_template.docx')

        # Передаем DAO объекты генератору
        generator.set_daos(setting_dao, author_dao)

        # Генерируем договоры только для сотрудников
        doc_buffers = generator.generate_for_all_authors(application, authors_data)

        if not doc_buffers:
            flash('Не удалось сгенерировать договоры', 'error')
            return redirect(f'/application/{application_id}')

        if len(doc_buffers) == 1:
            # Если только один сотрудник
            doc_buffer = doc_buffers[0]

            product_name = "product"
            if len(application) > 8 and application[8]:
                product_name = str(application[8]).strip()[:50]

            filename = f"dogovor_ob_otch_iskl_prava_{product_name}_{application_id}.docx"

            return send_file(
                doc_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=filename
            )
        else:
            # Если несколько сотрудников
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i, doc_buffer in enumerate(doc_buffers):
                    author_name = "author"

                    product_name = "product"
                    if len(application) > 8 and application[8]:
                        product_name = str(application[8]).strip()[:30]

                    filename = f"dogovor_ob_otch_iskl_prava_author_{i + 1}.docx"
                    zip_file.writestr(filename, doc_buffer.getvalue())

            zip_buffer.seek(0)

            product_name_short = "product"
            if len(application) > 8 and application[8]:
                product_name_short = str(application[8]).strip()[:30]

            zip_filename = f"dogovor_ob_otch_iskl_prava_{product_name_short}.zip"

            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name=zip_filename
            )

    except Exception as e:
        error_msg = str(e)
        print(f"ERROR in generate_employee_assignment_agreements: {error_msg}")
        import traceback
        traceback.print_exc()

        flash('Произошла ошибка при генерации договоров. Обратитесь к администратору', 'error')
        return redirect(f'/application/{application_id}')


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
def get_departments():
    """Получение списка отделов"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM departments ORDER BY name")
        return cur.fetchall()


def get_computer_types():
    """Получение списка типов компьютеров"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM computer_types ORDER BY name")
        return cur.fetchall()


def get_programming_languages():
    """Получение списка языков программирования"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM programming_languages ORDER BY name")
        return cur.fetchall()


def get_os_types():
    """Получение списка операционных систем"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM os_types ORDER BY name")
        return cur.fetchall()


def get_expert_roles():
    """Получение списка ролей экспертов"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM expert_roles ORDER BY name")
        return cur.fetchall()


def get_application_statuses():
    """Получение списка статусов заявок"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM application_statuses ORDER BY id")
        return cur.fetchall()


def get_application_status_by_id(status_id):
    """Получение статуса заявки по ID"""
    with db.get_cursor() as cur:
        cur.execute("SELECT id, name FROM application_statuses WHERE id = %s", (status_id,))
        result = cur.fetchone()
        if result:
            return {'id': result[0], 'name': result[1]}
        return None


def get_experts():
    """Получение списка экспертов"""
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM expert ORDER BY surname, name")
        return cur.fetchall()


if __name__ == '__main__':
    app.run(debug=True, port=5000)
