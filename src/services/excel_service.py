from typing import List

from src.core import config
from src.models import AuthorityModel
from src.utils import PathCreator, FileHandler


class ExcelService:
    def __init__(self):
        self._path_creator = PathCreator()
        self._file_handler = FileHandler()

    def read_excel(self) -> List[AuthorityModel]:

        sheet = self._file_handler.get_excel_data_sheet()

        row = config.excel_file_start_row
        authorities_data = []

        while True:

            authority_teryt = sheet[f'{config.excel_file_teryt_column}{row}'].value
            if not authority_teryt:
                break

            authority_data = AuthorityModel()
            authority_data.set_authority_teryt(authority_teryt)

            authority_data.set_authority_name(sheet[f'{config.excel_file_authority_name_column}{row}'].value)
            authority_data.set_governor_title(sheet[f'{config.excel_file_governor_title_column}{row}'].value)
            authority_data.set_governor_gender(sheet[f'{config.excel_file_governor_gender_column}{row}'].value)
            authority_data.set_governor_name(sheet[f'{config.excel_file_governor_name_column}{row}'].value)
            authority_data.set_authority_office_name(sheet[f'{config.excel_file_authority_office_name_column}{row}'].value)
            authority_data.set_authority_office_mail(sheet[f'{config.excel_file_authority_office_mail_column}{row}'].value)

            authorities_data.append(authority_data)
            row += 1

        print("Pomyślnie przeczytano wszystkie dane gmin.")

        return authorities_data
