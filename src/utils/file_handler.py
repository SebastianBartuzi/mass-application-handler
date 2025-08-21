import json
import os
from pathlib import Path

import openpyxl
import win32com.client as win32
from typing import Dict, Optional, Any, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx2pdf import convert as convert_docx_to_pdf
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from openpyxl.worksheet.worksheet import Worksheet
from win32com.client import CDispatch


class FileHandler:

    def set_excel_cell(self, sheet: Worksheet, cell_id: str, cell_value: Any):

        try:

            if cell_value is not None:

                if cell_value is True:
                    sheet[cell_id].value = 1

                elif cell_value is False:
                    sheet[cell_id].value = 0

                else:
                    sheet[cell_id].value = cell_value

        except Exception as e:
            raise Exception(f"Błąd podczas wypełniana komórki Excel {cell_id} z wartością {cell_value}: {e}")

    def get_excel_cell(self, sheet: Worksheet, cell_id: str) -> Any:

        try:

            return sheet[cell_id].value

        except Exception as e:
            raise Exception(f"Błąd podczas pobierania wartości komórki Excel {cell_id}: {e}")

    def get_file_extension(self, file_path: str) -> str:

        try:
            _, file_extension = os.path.splitext(file_path)
            return file_extension.lower()

        except Exception as e:
            raise Exception(f"Błąd podczas ekstrakcji rozszerzenia pliku {file_path}: {e}")

    def get_outlook_instance(self) -> CDispatch:

        try:
            return win32.Dispatch('outlook.application')

        except Exception as e:
            raise Exception(f"Błąd podczas pobierania instancji programu Outlook: {e}")

    def check_file_exists(self, file_path: str) -> bool:

        try:
            return os.path.exists(file_path)

        except Exception as e:
            raise Exception(f"Błąd podczas sprawdzania czy plik {file_path} istnieje: {e}")

    def get_excel_workbook_and_sheet(self, file_path: str, tab_name: str) -> Tuple[Workbook, Worksheet]:

        try:
            workbook = openpyxl.load_workbook(file_path)
            return workbook, workbook[tab_name]

        except Exception as e:
            raise Exception(f"Błąd podczas pobierania danych z Excel'a {file_path}: {e}")

    def save_workbook(self, workbook: Workbook, file_path: str, tab_name: str, start_cell: str, end_cell: str) -> None:

        try:
            sheet_to_table = workbook[tab_name]

            table_range = f"{start_cell}:{end_cell}"

            tab = Table(displayName="", ref=table_range)
            style = TableStyleInfo(
                showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False
            )
            tab.tableStyleInfo = style

            sheet_to_table.add_table(tab)
            print(f"Utworzono tabelę w arkuszu '{tab_name}' w zakresie '{table_range}'.")
            workbook.save(file_path)

        except Exception as e:
            raise Exception(f"Błąd podczas zapisywania danych do pliku Excel {file_path}: {e}")

    def save_docx(self, docx_data: Document, file_path: str) -> None:

        try:
            docx_data.save(file_path)

        except Exception as e:
            raise Exception(f"Błąd podczas zapisywania do ścieżki {file_path}: {e}")

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> None:

        try:
            convert_docx_to_pdf(docx_path, pdf_path)

        except Exception as e:
            raise Exception(f"Błąd podczas konwertowania ścieżki {docx_path} do ścieżki {pdf_path}: {e}")

    def read_string_from_txt(self, file_path: str) -> Optional[str]:

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            return content

        except Exception as e:
            raise Exception(f"Błąd podczas czytania pliku .txt {file_path}: {e}")

    def read_dict_from_json(self, file_path: str) -> Optional[Dict[str, Any]]:

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            return data

        except Exception as e:
            raise Exception(f"Błąd podczas czytania pliku .json {file_path}: {e}")

    def remove_file(self, file_path: str) -> None:

        try:
            os.remove(file_path)

        except Exception as e:
            raise Exception(f"Błąd podczas usuwania ścieżki {file_path}: {e}")

    def ensure_docx_formatting(self, paragraph: Paragraph, original_text: str, modified_text: str):

        try:
            if original_text != modified_text or len(paragraph.runs) > 1:

                first_run_format = {}

                if paragraph.runs:

                    first_run = paragraph.runs[0]
                    first_run_format = {
                        "bold": first_run.bold,
                        "italic": first_run.italic,
                        "underline": first_run.underline,
                        "size": first_run.font.size,
                        "color_rgb": first_run.font.color.rgb,
                        "name": first_run.font.name
                    }

                while len(paragraph.runs) > 0:

                    p_element = paragraph._element
                    r_element = paragraph.runs[0]._element
                    p_element.remove(r_element)

                new_run = paragraph.add_run(modified_text)

                if "bold" in first_run_format: new_run.bold = first_run_format["bold"]
                if "italic" in first_run_format: new_run.italic = first_run_format["italic"]
                if "underline" in first_run_format: new_run.underline = first_run_format["underline"]
                if "size" in first_run_format and first_run_format["size"]:
                    new_run.font.size = first_run_format["size"]
                if "color_rgb" in first_run_format and first_run_format["color_rgb"]:
                    new_run.font.color.rgb = first_run_format["color_rgb"]
                if "name" in first_run_format and first_run_format["name"]:
                    new_run.font.name = first_run_format["name"]

        except Exception as e:
            raise Exception(f"Błąd podczas zachowywania formatowania .docx: {e}")

    def convert_docx_to_txt(self, file_path: str) -> str:
        docx_path = Path(file_path)

        if not docx_path.exists():
            print(f"Error: The file '{file_path}' does not exist.")
            return file_path

        if docx_path.suffix.lower() != '.docx':
            return file_path

        print(f"Attempting to convert '{docx_path.name}' to text...")

        try:
            # 3. Open the .docx document
            document = Document(docx_path)

            # 4. Extract text content, paragraph by paragraph
            text_content = []
            for paragraph in document.paragraphs:
                # Append the text of the current paragraph, followed by a newline
                text_content.append(paragraph.text)

            # 5. Create the output file path with a .txt extension
            # `with_suffix` is a convenient pathlib method for changing the file extension
            txt_path = docx_path.with_suffix('.txt')

            # 6. Write the extracted text to the new .txt file
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                # Join all the paragraph texts with a newline
                txt_file.write('\n'.join(text_content))

            print(f"Successfully converted and saved text to '{txt_path}'.")

            self.remove_file(file_path)
            return str(txt_path)

        except Exception as e:
            print(f"An error occurred during conversion: {e}")
            return file_path

    def is_gemini_accepted_file(self, file_path: str) -> bool:

        try:
            accepted_extensions = {
                '.pdf', '.txt', '.xls', '.xlsx', '.csv', '.tsv',
                '.rtf', '.dot', '.dotx', '.hwp', '.hwpx', '.png', '.jpeg', '.jpg',
                '.webp', '.mp4', '.mov', '.webm', '.flv', '.mpeg', '.mpg', '.3gpp',
                '.mp3', '.wav', '.flac', '.aac', '.mpa', '.mpga', '.opus', '.pcm'
            }

            extensions_to_process = {'.docx'}

            _, file_extension = os.path.splitext(file_path)
            file_extension = file_extension.lower()

            if file_extension in accepted_extensions or file_extension in extensions_to_process:
                return True
            else:
                print(f"UWAGA! Nieakceptowalne rozszerzenie pliku: {file_extension} dla {file_path}. Plik nie zostanie "
                      f"uwzględniony w zapytaniu do AI.")

        except Exception as e:
            raise Exception(f"Błąd podczas sprawdzania czy rozszerzenie pliku {file_path} jest akceptowane przez "
                            f"Gemini: {e}")