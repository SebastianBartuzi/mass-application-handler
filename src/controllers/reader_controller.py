from src.services import MailService, PromptService, ExcelService
from src.utils import FileHandler


class ReaderController:
    def __init__(self):
        self.excel_service = ExcelService()
        self.mail_service = MailService()
        self.prompt_service = PromptService()
        self.file_handler = FileHandler()

    def build(self):
        outlook = self.file_handler.get_outlook_instance()

        authorities_data = self.excel_service.read_excel()
        mails_data = self.mail_service.read_emails_from_inbox(outlook)

        for mail_data in mails_data:
            prompt = self.prompt_service.get_mail_analysis_prompt(mail_data, authorities_data)