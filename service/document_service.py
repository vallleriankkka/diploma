import io
import os
from docx import Document
from datetime import datetime


class DocumentService:
    def __init__(self, application_dao, author_dao):
        self.application_dao = application_dao
        self.author_dao = author_dao

    def generate_abstract(self, application_id):
        """Генерация реферата"""
        application = self.application_dao.get_by_id(application_id)
        authors = self.application_dao.get_application_authors(application_id)

        if not application:
            return None

        # Создаем документ
        doc = Document()

        # ... (вся логика форматирования документа)

        return doc

    def generate_consent(self, application_id):
        """Генерация согласия"""
        # ... (логика генерации согласия)
        pass

    def generate_notification(self, application_id):
        """Генерация уведомления"""
        # ... (логика генерации уведомления)
        pass

    def save_to_stream(self, doc):
        """Сохраняем документ в поток"""
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        return file_stream
