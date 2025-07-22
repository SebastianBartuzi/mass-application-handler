import os

from typing import List
from dotenv import load_dotenv

from src.models import AuthorityModel
from src.utils import PathCreator, FileHandler, TypeConverter


class ExcelService:
    def __init__(self):
        load_dotenv()
        self.path_creator = PathCreator()
        self.type_converter = TypeConverter()
        self.file_handler = FileHandler()
        self.excel_file_start_row = self.type_converter.str_to_int(os.getenv('EXCEL_FILE_START_ROW'))
        self.excel_file_teryt_column = os.getenv('EXCEL_FILE_TERYT_COLUMN')
        self.excel_file_authority_name_column = os.getenv('EXCEL_FILE_AUTHORITY_NAME_COLUMN')
        self.excel_file_governor_title_column = os.getenv('EXCEL_FILE_GOVERNOR_TITLE_COLUMN')
        self.excel_file_governor_gender_column = os.getenv('EXCEL_FILE_GOVERNOR_GENDER_COLUMN')
        self.excel_file_governor_name_column = os.getenv('EXCEL_FILE_GOVERNOR_NAME_COLUMN')
        self.excel_file_authority_office_name_column = os.getenv('EXCEL_FILE_AUTHORITY_OFFICE_NAME_COLUMN')
        self.excel_file_authority_office_mail_column = os.getenv('EXCEL_FILE_AUTHORITY_OFFICE_MAIL_COLUMN')

    def read_excel(self) -> List[AuthorityModel]:

        sheet = self.file_handler.get_excel_data_sheet()

        row = self.excel_file_start_row
        authorities_data = []

        while True:

            authority_teryt = sheet[f'{self.excel_file_teryt_column}{row}'].value
            if not authority_teryt:
                break

            authority_data = AuthorityModel()
            authority_data.set_authority_teryt(authority_teryt)

            authority_data.set_authority_name(sheet[f'{self.excel_file_authority_name_column}{row}'].value)
            authority_data.set_governor_title(sheet[f'{self.excel_file_governor_title_column}{row}'].value)
            authority_data.set_governor_gender(sheet[f'{self.excel_file_governor_gender_column}{row}'].value)
            authority_data.set_governor_name(sheet[f'{self.excel_file_governor_name_column}{row}'].value)
            authority_data.set_authority_office_name(sheet[f'{self.excel_file_authority_office_name_column}{row}'].value)
            authority_data.set_authority_office_mail(sheet[f'{self.excel_file_authority_office_mail_column}{row}'].value)

            authorities_data.append(authority_data)
            row += 1

        print("Pomyślnie przeczytano wszystkie dane gmin.")

        return authorities_data
