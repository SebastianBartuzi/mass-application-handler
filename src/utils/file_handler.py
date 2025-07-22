import os
import openpyxl
import win32com.client as win32
from docx import Document
from docx.text.paragraph import Paragraph
from docx2pdf import convert as convert_docx_to_pdf

from dotenv import load_dotenv
from openpyxl.worksheet.worksheet import Worksheet
from win32com.client import CDispatch

from src.utils.path_creator import PathCreator


class FileHandler:
    def __init__(self) -> None:
        load_dotenv()
        self.path_creator = PathCreator()
        self.excel_file_tab_name = os.getenv('EXCEL_FILE_TAB_NAME')

    def get_outlook_instance(self) -> CDispatch:
        return win32.Dispatch('outlook.application')

    def check_file_exists(self, file_path: str) -> bool:
        return os.path.exists(file_path)

    def get_excel_data_sheet(self) -> Worksheet:
        try:
            workbook = openpyxl.load_workbook(self.path_creator.get_excel_file_path())
            return workbook[self.excel_file_tab_name]
        except Exception as e:
            raise Exception(f"Błąd podczas pobierania danych z Excel'a: {e}")

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

    def remove_file(self, file_path: str) -> None:
        try:
            os.remove(file_path)
        except Exception as e:
            raise Exception(f"Błąd podczas usuwania ścieżki {file_path}: {e}")

    def ensure_docx_formatting(self, paragraph: Paragraph, original_text: str, modified_text: str):
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