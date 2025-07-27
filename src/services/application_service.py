from docx import Document

from src.models import AuthorityModel
from src.enums import ApplicationPlaceholderEnum
from src.utils import PathCreator, Utils, FileHandler


class ApplicationService:

    def __init__(self) -> None:
        self._path_creator = PathCreator()
        self._file_handler = FileHandler()
        self._utils = Utils()

    def generate_application_pdf(self, authority_data: AuthorityModel) -> None:

        if not self._file_handler.check_file_exists(authority_data.get_application_pdf_path()):

            print(f"Tworzenie pliku PDF dla gminy o TERYT {authority_data.get_authority_teryt()}.")

            application_doc = Document(self._path_creator.get_application_template_path())

            replacements = {
                ApplicationPlaceholderEnum.DATE.value: self._utils.get_today_date(),
                ApplicationPlaceholderEnum.ADDRESSEE_NAME.value: authority_data.get_governor_name(),
                ApplicationPlaceholderEnum.ADDRESSEE_TITLE.value: authority_data.get_governor_title(),
                ApplicationPlaceholderEnum.AUTHORITY_OFFICE_NAME.value: authority_data.get_authority_office_name(),
                ApplicationPlaceholderEnum.SALUTATION.value: self._utils.get_salutation_denominator(
                    authority_data.get_authority_teryt(), authority_data.get_governor_gender()
                )
            }

            for paragraph in application_doc.paragraphs:
                original_text = paragraph.text
                modified_text = original_text

                for old_text, new_text in replacements.items():
                    modified_text = modified_text.replace(old_text, new_text)

                self._file_handler.ensure_docx_formatting(paragraph, original_text, modified_text)

            docx_path, pdf_path = authority_data.get_application_docx_path(), authority_data.get_application_pdf_path()
            self._file_handler.save_docx(application_doc, docx_path)
            self._file_handler.convert_docx_to_pdf(docx_path, pdf_path)
            self._file_handler.remove_file(docx_path)

            print(f"Pomyślnie wygenerowano PDF dla gminy o TERYT {authority_data.get_authority_teryt()}.")

        else:

            print(f"Plik PDF już istnieje dla gminy o TERYT {authority_data.get_authority_teryt()}.")