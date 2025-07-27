import random

from src.services import ApplicationService, ExcelService, MailService
from src.utils import FileHandler

class SenderController:
    def __init__(self) -> None:
        self.application_service = ApplicationService()
        self.excel_service = ExcelService()
        self.mail_service = MailService()
        self.file_handler = FileHandler()

    def build(self, test_mode: bool = True) -> None:
        outlook = self.file_handler.get_outlook_instance()

        authorities_data = self.excel_service.read_data_excel()

        if test_mode:
            authorities_data = random.sample(authorities_data, 5)

        for authority_data in authorities_data:
            self.application_service.generate_application_pdf(authority_data)
            self.mail_service.issue_authority_email(authority_data, outlook, test_mode)

        self.mail_service.issue_send_confirmation(outlook)