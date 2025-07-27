import os

from src.core import config

class PathCreator:
    def _get_absolute_temp_path(self) -> str:

        absolute_temp_path = os.path.abspath(config.temp_path)

        if not os.path.exists(absolute_temp_path):
            os.makedirs(absolute_temp_path)

        return absolute_temp_path

    def _get_attachments_dir_path(self) -> str:

        attachments_dir_path = os.path.abspath(config.attachments_dir_path)

        if not os.path.exists(attachments_dir_path):
            os.makedirs(attachments_dir_path)

        return attachments_dir_path

    def get_excel_file_path(self) -> str:
        return os.path.abspath(config.excel_file_path)

    def get_application_template_path(self) -> str:
        return os.path.abspath(config.application_template_path)

    def get_application_docx_path(self, teryt: str) -> str:
        return f"{self._get_absolute_temp_path()}/{teryt}.docx"

    def get_application_pdf_path(self, teryt: str) -> str:
        return f"{self._get_absolute_temp_path()}/{teryt}.pdf"

    def get_attachment_path(self, mail_id: int, attachment_id: int, file_extension: str) -> str:
        return f"{self._get_attachments_dir_path()}/{str(mail_id)}_{str(attachment_id)}{file_extension}"