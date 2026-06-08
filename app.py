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
import json


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
                'power_of_attorney_id': current_settings_tuple[1],
                'sum_id': current_settings_tuple[2],
                'provost_id': current_settings_tuple[3],
                'poa_number': current_settings_tuple[4],
                'poa_date': current_settings_tuple[5],
                'sum_value': current_settings_tuple[6],
                'provost_name': current_settings_tuple[7],
                'provost_post': current_settings_tuple[8]
            }

        # Загружаем опции для выпадающих списков
        power_of_attorney_options = setting_dao.get_reference_options('power_of_attorney')
        sum_options = setting_dao.get_reference_options('sum')
        provost_options = setting_dao.get_reference_options('provost')

        # Загружаем справочники для вкладки управления
        programming_languages = product_dao.get_all_programming_languages()
        os_types = product_dao.get_all_os_types()
        computer_types = product_dao.get_all_computer_types()

        # Загружаем роли авторов
        author_roles = application_dao.get_all_author_roles()

        # ЗАГРУЖАЕМ СПИСОК ПОЛЬЗОВАТЕЛЕЙ
        users = user_dao.get_all_users()

        return render_template('settings/index.html',
                               current_settings=current_settings,
                               power_of_attorney_options=power_of_attorney_options,
                               sum_options=sum_options,
                               provost_options=provost_options,
                               programming_languages=programming_languages,
                               os_types=os_types,
                               computer_types=computer_types,
                               author_roles=author_roles,
                               users=users)
    except Exception as e:
        flash(f'Ошибка загрузки настроек: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/admin/settings/save', methods=['POST'])
@login_required
@admin_required
def save_settings():
    """Сохранение настроек"""
    print("=== SAVE SETTINGS CALLED ===")
    print("Request method:", request.method)
    print("Content-Type:", request.headers.get('Content-Type'))
    print("is_json:", request.is_json)

    try:
        # Пробуем получить данные разными способами
        if request.is_json:
            data = request.get_json()
            print("Got JSON data:", data)
        else:
            data = request.form.to_dict()
            print("Got FORM data:", data)

        setting_data = {
            'power_of_attorney_id': data.get('power_of_attorney_id') or None,
            'sum_id': data.get('sum_id') or None,
            'provost_id': data.get('provost_id') or None
        }

        print("Saving setting_data:", setting_data)

        # Сохраняем настройки
        setting_id = setting_dao.save(setting_data)

        print("Saved successfully, setting_id:", setting_id)

        response = jsonify({
            'success': True,
            'message': 'Настройки успешно сохранены',
            'setting_id': setting_id
        })
        print("Returning response:", response.get_data(as_text=True))
        return response

    except Exception as e:
        import traceback
        print("ERROR:", str(e))
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/admin/settings/update_single', methods=['POST'])
@login_required
@admin_required
def update_single_setting():
    """Обновление одной настройки"""
    try:
        data = request.get_json()
        setting_type = data.get('setting_type')
        reference_id = data.get('reference_id')

        if not setting_type:
            return jsonify({'success': False, 'error': 'Не указан тип настройки'})

        # Преобразуем reference_id в int, если он есть
        ref_id = int(reference_id) if reference_id else None

        # Сопоставляем тип настройки с полем в таблице
        field_mapping = {
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


# ========== УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ ==========

@app.route('/admin/users/create', methods=['POST'])
@login_required
@admin_required
def admin_create_user():
    """Создание пользователя администратором"""
    try:
        data = request.get_json()
        username = data.get('username', '').strip()
        password = data.get('password', '')
        role = data.get('role', 'user')

        if not username:
            return jsonify({'success': False, 'error': 'Логин не может быть пустым'})

        if not password:
            return jsonify({'success': False, 'error': 'Пароль не может быть пустым'})

        if len(password) < 4:
            return jsonify({'success': False, 'error': 'Пароль должен содержать не менее 4 символов'})

        # Создаем пользователя
        success, result = user_dao.create_user_by_admin(username, password, role)

        if success:
            return jsonify({
                'success': True,
                'message': 'Пользователь успешно создан',
                'user': result
            })
        else:
            return jsonify({'success': False, 'error': result.get('error', 'Ошибка при создании пользователя')})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/users/update_role', methods=['POST'])
@login_required
@admin_required
def admin_update_user_role():
    """Обновление роли пользователя"""
    try:
        data = request.get_json()
        user_id = data.get('user_id')
        role = data.get('role')

        if not user_id:
            return jsonify({'success': False, 'error': 'Не указан ID пользователя'})

        if role not in ['user', 'admin']:
            return jsonify({'success': False, 'error': 'Неверная роль'})

        success = user_dao.update_user_role(user_id, role)

        if success:
            return jsonify({'success': True, 'message': 'Роль успешно обновлена'})
        else:
            return jsonify({'success': False, 'error': 'Пользователь не найден'})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/admin/users/delete', methods=['POST'])
@login_required
@admin_required
def admin_delete_user():
    """Удаление пользователя"""
    try:
        data = request.get_json()
        user_id = data.get('user_id')

        if not user_id:
            return jsonify({'success': False, 'error': 'Не указан ID пользователя'})

        # Нельзя удалить самого себя
        if user_id == session['user_id']:
            return jsonify({'success': False, 'error': 'Нельзя удалить самого себя'})

        success, message = user_dao.delete_user(user_id)

        if success:
            return jsonify({'success': True, 'message': message})
        else:
            return jsonify({'success': False, 'error': message})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ===== Роуты для создания новых записей в справочниках =====

@app.route('/admin/settings/create_power_of_attorney', methods=['POST'])
@login_required
@admin_required
def create_power_of_attorney():
    """Создание новой доверенности"""
    try:
        data = request.get_json()
        number = data.get('number', '').strip()
        date_of_receipt = data.get('date_of_receipt', '')

        if not number:
            return jsonify({'success': False, 'error': 'Номер доверенности не может быть пустым'})

        if not date_of_receipt:
            return jsonify({'success': False, 'error': 'Дата получения не может быть пустой'})

        # Проверяем, не существует ли уже такой доверенности
        if setting_dao.check_exists('power_of_attorney', 'number', number):
            return jsonify({'success': False, 'error': 'Доверенность с таким номером уже существует'})

        # Создаем новую запись
        poa_id = setting_dao.create_power_of_attorney(number, date_of_receipt)

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
        data = request.get_json()
        value_str = data.get('value', '').strip()

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
        data = request.get_json()
        full_name = data.get('full_name', '').strip()
        post = data.get('post', '').strip()

        if not full_name:
            return jsonify({'success': False, 'error': 'ФИО не может быть пустым'})

        # Проверяем, не существует ли уже такое ФИО
        if setting_dao.check_exists('provost', 'full_name', full_name):
            return jsonify({'success': False, 'error': 'Проректор с таким ФИО уже существует'})

        # Создаем новую запись
        provost_id = setting_dao.create_provost(full_name, post or None)

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

        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types',
            'author_role': 'expert_roles'
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
        if reference_type == 'author_role':
            with application_dao.db.get_cursor() as cur:
                cur.execute("INSERT INTO expert_roles (name) VALUES (%s) RETURNING id, name", (name,))
                result = cur.fetchone()
                reference = {'id': result[0], 'name': result[1]}
        else:
            reference = product_dao.create_reference(reference_type, name)

        return jsonify({
            'success': True,
            'message': 'Запись успешно создана',
            'reference': reference,
            'reference_type': reference_type
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

        table_mapping = {
            'programming_language': 'programming_languages',
            'os_type': 'os_types',
            'computer_type': 'computer_types',
            'author_role': 'expert_roles'
        }

        table_name = table_mapping.get(reference_type)
        if not table_name:
            return jsonify({'success': False, 'error': 'Неизвестный тип справочника'})

        # Проверка существования другого элемента с таким же названием
        with product_dao.db.get_cursor() as cur:
            cur.execute(f"SELECT COUNT(*) FROM {table_name} WHERE name = %s AND id != %s", (name, ref_id))
            if cur.fetchone()[0] > 0:
                return jsonify({'success': False, 'error': 'Такое название уже существует'})

        # Обновляем запись
        if reference_type == 'author_role':
            with application_dao.db.get_cursor() as cur:
                cur.execute("UPDATE expert_roles SET name = %s WHERE id = %s", (name, ref_id))
                success = cur.rowcount > 0
        else:
            success = product_dao.update_reference(reference_type, ref_id, name)

        if success:
            return jsonify({
                'success': True,
                'message': 'Запись успешно обновлена',
                'reference': {'id': ref_id, 'name': name},
                'reference_type': reference_type
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

        # Для ролей авторов проверяем, используется ли роль
        if reference_type == 'author_role':
            with application_dao.db.get_cursor() as cur:
                cur.execute("SELECT COUNT(*) FROM application_authors WHERE expert_role_id = %s", (ref_id,))
                if cur.fetchone()[0] > 0:
                    return jsonify({
                        'success': False,
                        'error': 'Невозможно удалить роль, так как она используется в заявках'
                    })
                cur.execute("DELETE FROM expert_roles WHERE id = %s", (ref_id,))
                success = cur.rowcount > 0
        else:
            success = product_dao.delete_reference(reference_type, ref_id)

        if success:
            return jsonify({'success': True, 'message': 'Запись успешно удалена'})
        else:
            return jsonify({'success': False, 'error': 'Запись не найдена'})

    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
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
            total_percent = sum(int(percent) for percent in percents if percent and percent.isdigit())
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

            if not application_dao.update_application(
                    id,
                    session['user_id'],
                    update_data,
                    session['role'] == 'admin'
            ):
                flash('Не удалось обновить заявку', 'error')
                return redirect(url_for('application_edit', id=id))

            # ========== ВАЖНО: ОБНОВЛЯЕМ АВТОРОВ ==========
            # Сначала удаляем всех существующих авторов
            current_authors = application_dao.get_application_authors(id)
            for author in current_authors:
                # author[0] - это id связи applicationauthor
                application_dao.remove_author_from_application(id, author[2])  # author[2] - author_id

            # Затем добавляем авторов заново
            author_ids = request.form.getlist('author_ids[]')
            role_ids = request.form.getlist('role_ids[]')
            percents_list = request.form.getlist('percents[]')
            times_list = request.form.getlist('times[]')
            descriptions_list = request.form.getlist('description[]')
            equipment_list = request.form.getlist('equipment[]')

            for i in range(len(author_ids)):
                author_id = author_ids[i]
                if not author_id:
                    continue

                # Проверяем, принадлежит ли автор пользователю
                owner_id = author_dao.get_by_user(int(author_id), session['user_id'])
                if owner_id or session['role'] == 'admin':
                    author_data = {
                        'author_id': int(author_id),
                        'role_id': int(role_ids[i]) if role_ids[i] else None,
                        'participation_percent': int(percents_list[i]) if percents_list[i] else 0,
                        'time_spent': int(times_list[i]) if i < len(times_list) and times_list[i] else None,
                        'description': descriptions_list[i] if i < len(descriptions_list) else '',
                        'equipment': equipment_list[i] if i < len(equipment_list) else '',
                    }
                    application_dao.add_author_to_application(id, author_data)

            flash('Заявка успешно обновлена!', 'success')
            return redirect(url_for('application_view', id=id))

        except Exception as e:
            flash(f'Ошибка при обновлении заявки: {str(e)}', 'error')
            import traceback
            traceback.print_exc()

    # GET запрос - показываем форму
    if session['role'] == 'admin':
        authors = author_dao.get_all()
        products = product_dao.get_all()
        commissions = commission_dao.get_all()
    else:
        authors = author_dao.get_all(session['user_id'])
        products = product_dao.get_all(session['user_id'])
        commissions = commission_dao.get_all(session['user_id'])

    application_authors = application_dao.get_application_authors(id)
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


@app.route('/commission/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def commission_edit(id):
    """Редактирование комиссии"""
    # Получаем комиссию
    commission = commission_dao.get_by_id(id)
    if not commission:
        flash('Комиссия не найдена', 'error')
        return redirect(url_for('commissions_list'))

    # Проверяем доступ (commission[4] - user_id)
    if session['role'] != 'admin' and commission[4] != session['user_id']:
        flash('У вас нет доступа к этой комиссии', 'error')
        return redirect(url_for('commissions_list'))

    if request.method == 'POST':
        try:
            # Получаем данные из формы
            name = request.form.get('name', '').strip()
            description = request.form.get('description', '').strip()

            if not name:
                flash('Название комиссии обязательно для заполнения', 'error')
                return redirect(url_for('commission_edit', id=id))

            # 1. Обновляем основную информацию о комиссии
            commission_dao.update(id, name, description)

            # 2. Обрабатываем удаление экспертов
            removed_experts_data = request.form.get('removed_experts_data', '[]')
            removed_experts = json.loads(removed_experts_data)

            for expert_id in removed_experts:
                commission_dao.remove_expert_from_commission(id, expert_id)

            # 3. Обрабатываем добавление существующих экспертов
            selected_experts_data = request.form.get('selected_experts_data', '[]')
            selected_experts = json.loads(selected_experts_data)

            for expert in selected_experts:
                commission_dao.add_expert_to_commission(
                    id,
                    expert['expert_id'],
                    expert['role_id']
                )

            # 4. Обрабатываем создание новых экспертов
            new_experts_data = request.form.get('new_experts_data', '[]')
            new_experts = json.loads(new_experts_data)

            for expert in new_experts:
                # Создаем эксперта через expert_dao
                expert_dict = {
                    'surname': expert['surname'],
                    'name': expert['name'],
                    'patronymic': expert.get('patronymic', ''),
                    'position': expert['position'],
                    'user_id': session['user_id']
                }
                expert_id = expert_dao.create(expert_dict)

                # Добавляем в комиссию
                commission_dao.add_expert_to_commission(
                    id,
                    expert_id,
                    expert['role_id']
                )

            flash('Комиссия успешно обновлена', 'success')
            return redirect(url_for('commission_view', id=id))

        except Exception as e:
            flash(f'Ошибка при обновлении комиссии: {str(e)}', 'error')
            return redirect(url_for('commission_edit', id=id))

    # GET запрос - показываем форму редактирования
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

    # Используем существующий метод get_commission_experts_detailed
    current_experts_raw = commission_dao.get_commission_experts_detailed(id)

    # Преобразуем в нужный формат (список словарей)
    current_experts = []
    for expert in current_experts_raw:
        current_experts.append({
            'expert_id': expert[0],  # id
            'surname': expert[1],  # surname
            'name': expert[2],  # name
            'patronymic': expert[3],  # patronymic
            'position': expert[4],  # position
            'role_name': expert[5],  # role_name
            'user_id': expert[6] if len(expert) > 6 else None
        })

    # Используем существующий метод get_available_experts_for_commission
    available_experts = commission_dao.get_available_experts_for_commission(id, session['user_id'])

    # Получаем список ролей
    roles_raw = commission_dao.get_commission_roles()
    commission_roles = []
    for role in roles_raw:
        commission_roles.append({
            'id': role[0],
            'name': role[1]
        })

    return render_template('commissions/edit.html',
                           commission=commission_dict,
                           current_experts=current_experts,
                           available_experts=available_experts,
                           commission_roles=commission_roles)


# ========== ГЕНЕРАЦИЯ ДОКУМЕНТОВ ==========
@app.route('/application/<int:id>/generate_consent')
def generate_consent(id):
    """Генерация документов согласия для всех авторов"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        print(f"DEBUG: Starting consent generation for application {id}")

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

        # Функция для очистки имени файла
        def clean_filename(text):
            if not text:
                return ""
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                text = text.replace(char, '')
            return text.strip()

        # Получаем список имен файлов
        doc_names = []
        for i, doc_buffer in enumerate(all_documents):
            if i < len(authors_data):
                author = authors_data[i]
                if len(author) > 7:
                    surname = clean_filename(str(author[7]).strip() if author[7] else '')
                    name = clean_filename(str(author[8]).strip() if author[8] else '')

                    if surname and name:
                        filename = f"sogl_na_ykaz_dan_{surname}_{name}.docx"
                    else:
                        filename = f"sogl_na_ykaz_dan_author_{i + 1}.docx"
                else:
                    filename = f"sogl_na_ykaz_dan_author_{i + 1}.docx"
            else:
                filename = f"sogl_na_ykaz_dan_author_{i + 1}.docx"
            doc_names.append(filename)

        # Если один документ - отправляем как файл
        if len(all_documents) == 1:
            single_doc = all_documents[0]
            single_doc.seek(0)
            return send_file(
                single_doc,
                as_attachment=True,
                download_name=doc_names[0],
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

        # Если несколько документов - создаем ZIP архив
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for doc_buffer, doc_name in zip(all_documents, doc_names):
                doc_buffer.seek(0)
                zip_file.writestr(doc_name, doc_buffer.getvalue())

        zip_buffer.seek(0)
        zip_filename = f"sogl_na_ykaz_dan_zayavka_{id}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации согласий: {e}", exc_info=True)
        flash(f'Ошибка генерации документов: {str(e)}', 'error')
        return redirect(url_for('application_view', id=id))


@app.route('/application/<int:application_id>/generate_abstract')
def generate_abstract_from_application(application_id):
    """Генерация аннотации из заявки (всегда один документ)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем ID продукта из заявки
        product_id = application[2]

        # Получаем данные продукта
        product_data = product_dao.get_by_id(product_id)
        if not product_data:
            flash('Продукт не найден', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов
        authors_data = application_dao.get_application_authors_detailed(application_id)

        # Создаем генератор
        generator = AbstractGenerator()
        doc = generator.generate(product_data, authors_data)

        if doc:
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            filename = f"referat_zayavka_{application_id}.docx"

            # ИЗМЕНЕНИЕ ЗДЕСЬ - используем Response вместо send_file
            from flask import Response
            response = Response(
                file_stream.getvalue(),
                mimetype='application/octet-stream',
                headers={
                    'Content-Disposition': f'attachment; filename="{filename}"',
                    'Content-Type': 'application/octet-stream',
                    'X-Content-Type-Options': 'nosniff'
                }
            )
            return response
        else:
            flash('Не удалось сгенерировать документ', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации реферата: {e}", exc_info=True)
        flash(f'Ошибка генерации документа: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_sogl')
def generate_sogl_from_application(application_id):
    """Генерация согласия на обработку персональных данных"""
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

        # Получаем детальные данные авторов для имен файлов
        authors_data = application_dao.get_application_authors_detailed(application_id)

        # Функция для очистки имени файла
        def clean_filename(text):
            if not text:
                return ""
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                text = text.replace(char, '')
            return text.strip()

        # Если один документ
        if len(documents) == 1:
            file_stream = io.BytesIO()
            documents[0].save(file_stream)
            file_stream.seek(0)

            # Формируем имя файла с ФИО автора
            filename = f"sogl_na_obr_pd_{application_id}.docx"
            if authors_data and len(authors_data) > 0:
                author = authors_data[0]
                if len(author) > 7:
                    surname = clean_filename(str(author[7]).strip() if author[7] else '')
                    name = clean_filename(str(author[8]).strip() if author[8] else '')
                    if surname and name:
                        filename = f"sogl_na_obr_pd_{surname}_{name}.docx"

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

        # Если несколько документов - ZIP архив
        zip_stream = io.BytesIO()
        with zipfile.ZipFile(zip_stream, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, doc in enumerate(documents):
                doc_stream = io.BytesIO()
                doc.save(doc_stream)
                doc_stream.seek(0)

                # Формируем имя файла с ФИО автора
                filename = f"sogl_na_obr_pd_author_{i + 1}.docx"
                if authors_data and i < len(authors_data):
                    author = authors_data[i]
                    if len(author) > 7:
                        surname = clean_filename(str(author[7]).strip() if author[7] else '')
                        name = clean_filename(str(author[8]).strip() if author[8] else '')
                        if surname and name:
                            filename = f"sogl_na_obr_pd_{surname}_{name}.docx"

                zip_file.writestr(filename, doc_stream.read())

        zip_stream.seek(0)
        filename = f"sogl_na_obr_pd_{application_id}.zip"

        return send_file(
            zip_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/zip'
        )

    except Exception as e:
        print(f"Ошибка при генерации согласия: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Ошибка при генерации согласия: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_conclusion')
def generate_conclusion_from_application(application_id):
    """Генерация заключения о возможности открытого опубликования (всегда один документ)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        user_id = session['user_id']

        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('applications_list'))

        # Поиск commission_id
        commission_id = None
        with db.get_cursor() as cur:
            cur.execute("""
                SELECT column_name 
                FROM information_schema.columns 
                WHERE table_name = 'application' 
                ORDER BY ordinal_position
            """)
            columns = [col[0] for col in cur.fetchall()]

            for col in columns:
                if 'commission' in col.lower():
                    cur.execute(f"SELECT {col} FROM application WHERE id = %s", (application_id,))
                    result = cur.fetchone()
                    if result and result[0]:
                        commission_id = result[0]
                        break

        if not commission_id:
            user_commissions = commission_dao.get_all(user_id)
            if user_commissions and len(user_commissions) > 0:
                commission_id = user_commissions[0][0]

        commission_experts = None
        if commission_id:
            commission_experts = commission_dao.get_commission_experts(commission_id)

            if commission_experts and len(commission_experts) > 0 and len(commission_experts[0]) < 6:
                with db.get_cursor() as cur:
                    cur.execute("""
                        SELECT 
                            e.id, e.surname, e.name, e.patronymic, e.position, cr.name as role_name
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
                    if detailed_experts:
                        commission_experts = detailed_experts

        # Получаем настройки
        settings_data = setting_dao.get_current()

        # Получаем авторов
        authors_data = application_dao.get_application_authors_detailed(application_id)

        # Создаем генератор
        generator = ConclusionGenerator()

        # Если нет данных комиссии, создаем тестовые данные
        if not commission_experts:
            test_experts = [
                (1, 'Иванов', 'Иван', 'Иванович', 'профессор', 'председатель'),
                (2, 'Петрова', 'Мария', 'Сергеевна', 'доцент', 'эксперт'),
                (3, 'Сидоров', 'Алексей', 'Петрович', 'старший преподаватель', 'эксперт')
            ]
            commission_experts = test_experts

        # Генерируем документ
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
            product_name = ""
            if len(application) > 8 and application[8]:
                product_name = str(application[8]).strip()

            safe_product_name = "".join(c for c in product_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            if not safe_product_name:
                safe_product_name = f"zayavka_{application_id}"
            if len(safe_product_name) > 50:
                safe_product_name = safe_product_name[:50]

            filename = f"zakl_o_vozm_otkr_opybl_{safe_product_name}.docx"

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('Не удалось сгенерировать документ заключения', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации заключения: {e}", exc_info=True)
        flash(f'Ошибка генерации документа: {str(e)}', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_distribution_agreement')
def generate_distribution_agreement(application_id):
    """Генерация соглашения о распределении вознаграждения (всегда один документ)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Получаем авторов
        authors_data = application_dao.get_application_authors_detailed(application_id)
        if not authors_data:
            flash('У заявки нет авторов', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Создаем генератор
        generator = DistributionAgreementGenerator(template_name='distributionagreement_generator.docx')
        doc = generator.generate(application, authors_data)

        if doc:
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            filename = f"sogl_o_raspr_vozn_{application_id}.docx"

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
        print(f"ERROR in generate_distribution_agreement: {e}")
        import traceback
        traceback.print_exc()
        flash('Произошла ошибка при генерации документа', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_rid_notification')
def generate_rid_notification(application_id):
    """Генерация уведомления о создании РИД (всегда один документ)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(url_for('application_view', id=application_id))

        # Создаем генератор
        generator = RIDNotificationGenerator(template_name='notification_template.docx')
        if hasattr(generator, 'set_daos'):
            generator.set_daos(setting_dao, application_dao)

        # Генерируем документ
        doc = generator.generate(application_id, session['user_id'])

        if doc:
            file_stream = io.BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)

            filename = f"yved_rid_{application_id}.docx"

            return send_file(
                file_stream,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            flash('Не удалось сгенерировать уведомление', 'error')
            return redirect(url_for('application_view', id=application_id))

    except Exception as e:
        print(f"ERROR in generate_rid_notification: {e}")
        import traceback
        traceback.print_exc()
        flash('Произошла ошибка при генерации уведомления', 'error')
        return redirect(url_for('application_view', id=application_id))


@app.route('/application/<int:application_id>/generate_employee_assignment_agreements')
def generate_assignment_agreement(application_id):
    """Генерация договоров об отчуждении исключительного права (для сотрудников)"""
    try:
        # Получаем заявку
        application = application_dao.get_by_id(application_id)
        if not application:
            flash('Заявка не найдена', 'error')
            return redirect(f'/application/{application_id}')

        # Получаем авторов
        authors_data = application_dao.get_application_authors_detailed(application_id)
        if not authors_data:
            flash('У заявки нет авторов', 'error')
            return redirect(f'/application/{application_id}')

        # Создаем генератор
        generator = AssignmentAgreementGenerator(template_name='assignment_agreement_template.docx')
        generator.set_daos(setting_dao, author_dao)

        # Генерируем договоры
        doc_buffers = generator.generate_for_all_authors(application, authors_data)

        if not doc_buffers:
            flash('Не удалось сгенерировать договоры', 'error')
            return redirect(f'/application/{application_id}')

        # Функция для очистки имени файла
        def clean_filename(text):
            if not text:
                return ""
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                text = text.replace(char, '')
            return text.strip()

        # Получаем название продукта
        product_name = "product"
        if len(application) > 8 and application[8]:
            product_name = clean_filename(str(application[8]).strip()[:50])

        # Если один договор - отправляем как файл
        if len(doc_buffers) == 1:
            doc_buffer = doc_buffers[0]
            doc_buffer.seek(0)

            # Формируем имя с ФИО автора
            filename = f"dogovor_otch_iskl_prava_{product_name}.docx"
            if authors_data and len(authors_data) > 0:
                author = authors_data[0]
                if len(author) > 7:
                    surname = clean_filename(str(author[7]).strip() if author[7] else '')
                    name = clean_filename(str(author[8]).strip() if author[8] else '')
                    if surname and name:
                        filename = f"dogovor_otch_iskl_prava_{surname}_{name}.docx"

            return send_file(
                doc_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=filename
            )

        # Если несколько договоров - ZIP архив
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, doc_buffer in enumerate(doc_buffers):
                doc_buffer.seek(0)

                # Формируем имя файла с ФИО автора
                filename = f"dogovor_otch_iskl_prava_author_{i + 1}.docx"
                if authors_data and i < len(authors_data):
                    author = authors_data[i]
                    if len(author) > 7:
                        surname = clean_filename(str(author[7]).strip() if author[7] else '')
                        name = clean_filename(str(author[8]).strip() if author[8] else '')
                        if surname and name:
                            filename = f"dogovor_otch_iskl_prava_{surname}_{name}.docx"

                zip_file.writestr(filename, doc_buffer.getvalue())

        zip_buffer.seek(0)
        zip_filename = f"dogovor_otch_iskl_prava_{product_name}.zip"

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_filename
        )

    except Exception as e:
        print(f"ERROR in generate_assignment_agreement: {e}")
        import traceback
        traceback.print_exc()
        flash('Произошла ошибка при генерации договоров', 'error')
        return redirect(f'/application/{application_id}')


# ========== ГЕНЕРАЦИЯ ВСЕХ ДОКУМЕНТОВ ОДНИМ АРХИВОМ ==========
@app.route('/application/<int:id>/generate_all_documents')
@login_required
def generate_all_documents(id):
    """Генерация всех документов для заявки одним архивом или файлом"""
    if 'user_id' not in session:
        return redirect(url_for('login'))

    try:
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

        # Получаем данные продукта
        product_id = application[2]
        product_data = product_dao.get_by_id(product_id)

        # Создаем буферы для всех документов
        all_documents = []
        doc_names = []

        # 1. Согласие на указание сведений (для каждого автора отдельно)
        consent_gen = ConsentGenerator()
        consent_docs = consent_gen.generate_all_authors(application, authors_data, settings_data)
        for i, doc in enumerate(consent_docs):
            author = authors_data[i] if i < len(authors_data) else None
            if author and len(author) > 7:
                filename = f"1_sogl_na_ykaz_dan_{author[7]}_{author[8]}.docx"
            else:
                filename = f"1_sogl_na_ykaz_dan_author_{i + 1}.docx"
            all_documents.append(doc)
            doc_names.append(filename)

        # 2. Согласие на обработку персональных данных (для каждого автора отдельно)
        sogl_gen = SoglGenerator()
        sogl_docs = sogl_gen.generate_for_each_author(application,
                                                      application_dao.get_application_authors(id), author_dao)
        for i, doc in enumerate(sogl_docs):
            author = authors_data[i] if i < len(authors_data) else None
            if author and len(author) > 7:
                filename = f"2_sogl_na_obr_pd_{author[7]}_{author[8]}.docx"
            else:
                filename = f"2_sogl_na_obr_pd_author_{i + 1}.docx"
            # Сохраняем doc в буфер
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            all_documents.append(doc_buffer)
            doc_names.append(filename)

        # 3. Реферат
        abstract_gen = AbstractGenerator()
        abstract_doc = abstract_gen.generate(product_data, authors_data)
        if abstract_doc:
            abstract_buffer = io.BytesIO()
            abstract_doc.save(abstract_buffer)
            abstract_buffer.seek(0)
            all_documents.append(abstract_buffer)
            doc_names.append("3_referat.docx")

        # 4. Уведомление о создании РИД
        rid_gen = RIDNotificationGenerator(template_name='notification_template.docx')
        if hasattr(rid_gen, 'set_daos'):
            rid_gen.set_daos(setting_dao, application_dao)
        rid_doc = rid_gen.generate(id, session['user_id'])
        if rid_doc:
            rid_buffer = io.BytesIO()
            rid_doc.save(rid_buffer)
            rid_buffer.seek(0)
            all_documents.append(rid_buffer)
            doc_names.append("4_yved_rid.docx")

        # 5. Соглашение о распределении вознаграждения
        dist_gen = DistributionAgreementGenerator(template_name='distributionagreement_generator.docx')
        dist_doc = dist_gen.generate(application, authors_data)
        if dist_doc:
            dist_buffer = io.BytesIO()
            dist_doc.save(dist_buffer)
            dist_buffer.seek(0)
            all_documents.append(dist_buffer)
            doc_names.append("5_sogl_o_raspr_vozn.docx")

        # 6. Заключение о возможности открытого опубликования
        conclusion_gen = ConclusionGenerator()
        commission_experts = None
        commission_id = None
        # Получаем commission_id из заявки
        for idx in [12, 13, 14, 15, 16, 17]:
            if len(application) > idx and application[idx]:
                if str(application[idx]).isdigit():
                    commission_id = int(application[idx])
                    break
        if commission_id:
            commission_experts = commission_dao.get_commission_experts(commission_id)
        conclusion_doc = conclusion_gen.generate(application, commission_experts, settings_data, authors_data)
        if conclusion_doc:
            conclusion_buffer = io.BytesIO()
            conclusion_doc.save(conclusion_buffer)
            conclusion_buffer.seek(0)
            all_documents.append(conclusion_buffer)
            doc_names.append("6_zakl_o_vozm_otkr_opybl.docx")

        # 7. Договор об отчуждении (только для сотрудников)
        assign_gen = AssignmentAgreementGenerator(template_name='assignment_agreement_template.docx')
        assign_gen.set_daos(setting_dao, author_dao)
        assign_buffers = assign_gen.generate_for_all_authors(application, authors_data)
        for i, buffer in enumerate(assign_buffers):
            author = authors_data[i] if i < len(authors_data) else None
            if author and len(author) > 7:
                filename = f"7_dog_otch_iskl_prava_{author[7]}_{author[8]}.docx"
            else:
                filename = f"7_dog_otch_iskl_prava_author_{i + 1}.docx"
            all_documents.append(buffer)
            doc_names.append(filename)

        if not all_documents:
            flash('Не удалось сгенерировать ни одного документа', 'error')
            return redirect(url_for('application_view', id=id))

        # Проверяем количество документов
        if len(all_documents) == 1:
            # Один документ - отправляем как файл
            single_doc = all_documents[0]
            single_doc.seek(0)
            return send_file(
                single_doc,
                as_attachment=True,
                download_name=doc_names[0],
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        else:
            # Несколько документов - создаем ZIP архив
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for doc_buffer, doc_name in zip(all_documents, doc_names):
                    doc_buffer.seek(0)
                    zip_file.writestr(doc_name, doc_buffer.getvalue())

            zip_buffer.seek(0)

            # Имя архива
            product_name = ""
            if len(application) > 8 and application[8]:
                product_name = application[8][:50].replace(' ', '_')
            else:
                product_name = f"zayavka_{id}"

            zip_filename = f"vse_dokumenty_{product_name}.zip"

            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name=zip_filename,
                mimetype='application/zip'
            )

    except Exception as e:
        current_app.logger.error(f"Ошибка генерации всех документов: {e}", exc_info=True)
        flash(f'Ошибка генерации документов: {str(e)}', 'error')
        return redirect(url_for('application_view', id=id))


@app.route('/download/documents_list')
@login_required
def download_documents_list():
    # Путь к файлу
    file_path = os.path.join('templates', 'documents', 'Perechen-neobkhodimykh-dokumentov.docx')

    # Просто отправляем файл на скачивание
    return send_file(
        file_path,
        as_attachment=True,  # как вложение (откроется диалог сохранения)
        download_name='Perechen-neobkhodimykh-dokumentov.docx',  # имя файла при скачивании
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

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
    """Получение списка ролей """
    with db.get_cursor() as cur:
        cur.execute("SELECT * FROM author_roles ORDER BY name")
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
    app.run(debug=True, host='0.0.0.0', port=5000)

