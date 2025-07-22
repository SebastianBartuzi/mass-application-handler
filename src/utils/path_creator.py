import os
from dotenv import load_dotenv

class PathCreator:
    def __init__(self) -> None:
        load_dotenv()
        self.excel_file_path = os.getenv('EXCEL_FILE_PATH')
        self.application_template_path = os.getenv('APPLICATION_TEMPLATE_PATH')
        self.temp_path = os.getenv('TMP_PATH')

    def get_excel_file_path(self) -> str:
        return os.path.abspath(self.excel_file_path)

    def get_application_template_path(self) -> str:
        return os.path.abspath(self.application_template_path)

    def _get_absolute_temp_path(self) -> str:

        absolute_temp_path = os.path.abspath(self.temp_path)
        if not os.path.exists(absolute_temp_path):
            os.makedirs(absolute_temp_path)

        return absolute_temp_path

    def get_application_docx_path(self, teryt: str) -> str:
        return f"{self._get_absolute_temp_path()}/{teryt}.docx"

    def get_application_pdf_path(self, teryt: str) -> str:
        return f"{self._get_absolute_temp_path()}/{teryt}.pdf"