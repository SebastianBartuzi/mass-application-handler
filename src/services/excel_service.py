from typing import List, Optional, Any

from openpyxl.worksheet.worksheet import Worksheet

from src.core import config
from src.models import AuthorityModel, ResponseMailModel
from src.utils import PathCreator, FileHandler, Utils


class ExcelService:
    def __init__(self):
        self._path_creator = PathCreator()
        self._file_handler = FileHandler()
        self._utils = Utils()

    def read_data_excel(self) -> List[AuthorityModel]:

        _, sheet = self._file_handler.get_excel_workbook_and_sheet(
            self._path_creator.get_data_excel_file_path(), config.data_excel_tab_name
        )

        row = config.data_excel_start_row
        authorities_data = []

        while True:

            authority_teryt = self._file_handler.get_excel_cell(sheet, f'{config.data_excel_teryt_column}{row}')
            if not authority_teryt:
                break

            authority_data = AuthorityModel()
            authority_data.set_authority_teryt(authority_teryt)

            authority_data.set_authority_name(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_authority_name_column}{row}')
            )
            authority_data.set_governor_title(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_governor_title_column}{row}')
            )
            authority_data.set_governor_gender(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_governor_gender_column}{row}')
            )
            authority_data.set_governor_name(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_governor_name_column}{row}')
            )
            authority_data.set_authority_office_name(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_authority_office_name_column}{row}')
            )
            authority_data.set_authority_office_mail(
                self._file_handler.get_excel_cell(sheet, f'{config.data_excel_authority_office_mail_column}{row}')
            )

            authorities_data.append(authority_data)
            row += 1

        print("Pomyślnie przeczytano wszystkie dane gmin.")

        return authorities_data

    def _find_authority_row_by_teryt(self, sheet: Worksheet, teryt: str) -> Optional[int]:

        for cell in sheet[config.report_excel_teryt_column]:
            if cell.value == teryt:
                return cell.row

        return None

    def fill_report_excel(self, mails_data: List[ResponseMailModel]) -> None:

        workbook, sheet = self._file_handler.get_excel_workbook_and_sheet(
            self._path_creator.get_report_template_excel_file_path(), config.report_excel_tab_name
        )

        for mail_data in mails_data:

            if (not mail_data.get_no_teryt_matched() and not mail_data.get_multiple_teryts_matched()
                and not mail_data.get_automatic_response()):

                teryt = mail_data.get_teryt()
                row = self._find_authority_row_by_teryt(sheet, teryt)

                if not row:
                    print(f"Nie znaleziono rzędu dla TERYT: {teryt}!")
                    continue

                self._file_handler.set_excel_cell(
                    sheet, f'{config.report_excel_last_answer_date_column}{row}', mail_data.get_mail_date()
                )

                attention_information = mail_data.get_attention_information()
                self._file_handler.set_excel_cell(
                    sheet, f'{config.report_excel_needs_attention_column}{row}', attention_information
                )
                if attention_information != "":
                    self._file_handler.set_excel_cell(
                        sheet,
                        f'{config.report_excel_additional_context_column}{row}',
                        mail_data.get_additional_info_response_type()
                    )

                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_offers_internships_column}{row}',
                    mail_data.get_offers_internships()
                )
                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_are_internships_paid_column}{row}',
                    mail_data.get_are_internships_paid()
                )
                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_plans_paid_internships_column}{row}',
                    mail_data.get_plans_paid_internships()
                )
                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_paid_internships_number_column}{row}',
                    mail_data.get_paid_internships_number()
                )
                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_internships_salaries_column}{row}',
                    mail_data.get_internships_salaries()
                )
                self._file_handler.set_excel_cell(
                    sheet,
                    f'{config.report_excel_additional_info_column}{row}',
                    mail_data.get_additional_info_questions_responses()
                )

                print(f"Pomyślnie wypełniono rząd tabeli dla TERYT: {teryt}.")

        self._file_handler.save_workbook(
            workbook, self._path_creator.get_report_excel_file_path(), config.report_excel_tab_name,
            config.report_excel_start_cell, config.report_excel_end_cell
        )
        print("Pomyślnie wypełniono raport!")