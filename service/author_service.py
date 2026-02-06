from datetime import date
import re


class AuthorService:
    def __init__(self, author_dao):
        self.author_dao = author_dao

    def calculate_age(self, birth_date, reference_date=None):
        """Рассчитать возраст на указанную дату (по умолчанию - сегодня)"""
        if reference_date is None:
            reference_date = date.today()

        return reference_date.year - birth_date.year - (
                (reference_date.month, reference_date.day) <
                (birth_date.month, birth_date.day)
        )

    def validate_author(self, author_data, author_id=None):
        errors = []

        # Проверка обязательных полей (добавьте это)
        required_fields = ['surname', 'name', 'birth_date', 'address',
                           'passport', 'snils', 'inn']

        for field in required_fields:
            if not author_data.get(field):
                field_names = {
                    'surname': 'Фамилия',
                    'name': 'Имя',
                    'birth_date': 'Дата рождения',
                    'address': 'Адрес',
                    'passport': 'Паспорт',
                    'snils': 'СНИЛС',
                    'inn': 'ИНН'
                }
                errors.append(f'Поле "{field_names.get(field, field)}" обязательно для заполнения')

        # Проверка возраста на текущую дату
        birth_date = author_data.get('birth_date')
        if birth_date:
            age = self.calculate_age(birth_date)
            if age < 14:
                errors.append('Автор должен быть старше 14 лет')

        # Проверка возраста на дату выдачи паспорта
        passport_issue_date = author_data.get('passport_issue_date')
        if birth_date and passport_issue_date:
            age_at_issue = self.calculate_age(birth_date, passport_issue_date)
            if age_at_issue < 14:
                errors.append('Паспорт не может быть выдан лицу младше 14 лет')

        # Проверка даты выдачи паспорта
        if passport_issue_date and passport_issue_date > date.today():
            errors.append('Дата выдачи паспорта не может быть в будущем')

        # Проверка уникальности паспорта
        passport = author_data.get('passport')
        if passport:
            if self.author_dao.check_unique('passport', passport, author_id):
                errors.append('Автор с такими паспортными данными уже существует')

        # Проверка уникальности СНИЛС
        snils = author_data.get('snils')
        if snils:
            if self.author_dao.check_unique('snils', snils, author_id):
                errors.append('Автор с таким СНИЛС уже существует')

        # Проверка уникальности ИНН
        inn = author_data.get('inn')
        if inn:
            if self.author_dao.check_unique('inn', inn, author_id):
                errors.append('Автор с таким ИНН уже существует')

        # Валидация форматов
        if passport and not self.is_valid_passport(passport):
            errors.append('Неверный формат паспортных данных (требуется 10 цифр)')

        if snils and not self.is_valid_snils(snils):
            errors.append('Неверный формат СНИЛС (требуется формат: XXX-XXX-XXX YY)')

        if inn and not self.is_valid_inn(inn):
            errors.append('Неверный формат ИНН (требуется 12 цифр)')

        # Проверка для студента (добавьте это)
        if author_data.get('is_student') and not author_data.get('education_level_id'):
            errors.append('Для студента необходимо указать уровень образования')

        return errors

    @staticmethod
    def is_valid_passport(passport):
        pattern = r'^\d{10}$'
        return bool(re.match(pattern, passport))

    @staticmethod
    def is_valid_snils(snils):
        pattern = r'^\d{3}-\d{3}-\d{3} \d{2}$'
        return bool(re.match(pattern, snils))

    @staticmethod
    def is_valid_inn(inn):
        pattern = r'^\d{12}$'
        return bool(re.match(pattern, inn))

    def can_access_author(self, author_id, user_id, user_role):
        owner_id = self.author_dao.get_by_user(author_id, user_id)
        if not owner_id:
            return False
        return user_role == 'admin' or owner_id == user_id
