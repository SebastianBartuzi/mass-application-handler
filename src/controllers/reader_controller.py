from src.services import MailService, PromptService, ExcelService, GeminiService, ResponseService
from src.utils import FileHandler


class ReaderController:
    def __init__(self):
        self._excel_service = ExcelService()
        self._mail_service = MailService()
        self._prompt_service = PromptService()
        self._gemini_service = GeminiService()
        self._response_service = ResponseService()
        self._file_handler = FileHandler()

    def build(self) -> None:

        outlook = self._file_handler.get_outlook_instance()

        authorities_data = self._excel_service.read_data_excel()
        mails_data = self._mail_service.read_new_emails_from_inbox(outlook)

        for mail_data in mails_data:

            prompt = self._prompt_service.get_mail_analysis_prompt(mail_data)
            response_data = self._gemini_service.analyse_mail(prompt, mail_data)

            self._response_service.extract_mail_analysis_response_data(response_data, mail_data)

            if not mail_data.get_teryt() and not mail_data.get_automatic_response():

                prompt = self._prompt_service.get_teryt_matcher_prompt(mail_data, authorities_data)
                response_data = self._gemini_service.match_teryt(prompt, mail_data)

                self._response_service.set_matched_teryt(response_data, mail_data)

            for attachment_path in mail_data.get_attachments_paths():
                self._file_handler.remove_file(attachment_path)

        self._excel_service.fill_report_excel(mails_data)

        self._mail_service.forward_attention_needed_mails(outlook, mails_data)
        self._mail_service.issue_read_confirmation(outlook, mails_data)
        self._mail_service.issue_stop_flag(outlook)
