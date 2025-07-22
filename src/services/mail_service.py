import os

from dotenv import load_dotenv
from win32com.client import CDispatch

from src.models import AuthorityModel
from src.utils import Utils, FileHandler, TypeConverter


class MailService:
    def __init__(self) -> None:
        self.file_handler = FileHandler()
        self.type_converter = TypeConverter()
        self.utils = Utils()
        load_dotenv()
        self.sender_email = os.getenv('SENDER_EMAIL')
        self.confirmation_addressees = os.getenv('CONFIRMATION_ADDRESSEES')
        self.successful_teryts = {}
        self.unsuccessful_teryts = {}

    def _get_authority_mail_title(self, authority_data: AuthorityModel) -> str:
        return f'Wniosek o udostępnienie informacji publicznej ({authority_data.get_authority_teryt()})'

    def _get_authority_mail_content(self, authority_data: AuthorityModel) -> str:
        salutation = self.utils.get_salutation_vocative(
            authority_data.get_authority_teryt(), authority_data.get_governor_gender()
        )
        return (f"{salutation},\nW załączniku przesyłam wniosek o udostępnienie informacji publicznej.\nW tytule "
                f"odpowiedzi zwrotnej proszę o zawarcie kodu TERYT gminy: {authority_data.get_authority_teryt()}.\n\n"
                f"Z poważaniem\nSebastian Bartuzi\nCzłonek Stowarzyszenia Młoda Lewica RP")

    def send_authority_email(self, authority_data: AuthorityModel, outlook: CDispatch, test_mode: bool) -> None:

        if test_mode:
            mail_to = self.confirmation_addressees
        else:
            mail_to = authority_data.get_authority_office_mail()

        try:

            mail = outlook.CreateItem(0)

            mail.To = mail_to
            mail.Subject = self._get_authority_mail_title(authority_data)
            mail.Body = self._get_authority_mail_content(authority_data)

            mail.SentOnBehalfOfName = self.sender_email

            attachment_path = authority_data.get_application_pdf_path()
            mail.Attachments.Add(attachment_path)

            mail.Send()
            print(f"E-mail do gminy o TERYT {authority_data.get_authority_teryt()} wysłano pomyślnie.")
            self.successful_teryts.update({authority_data.get_authority_teryt(): mail_to})

            self.file_handler.remove_file(attachment_path)

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a do gminy o TERYT {authority_data.get_authority_teryt()}: {e}")
            self.unsuccessful_teryts.update({authority_data.get_authority_teryt(): (mail_to, e)})

    def _get_confirmation_mail_title(self) -> str:
        return f"mlAI: Potwierdzenie operacji z dnia {self.utils.get_today_date()}"

    def _get_confirmation_mail_content(self) -> str:
        successful_sends_len = len(self.successful_teryts)
        unsuccessful_sends_len = len(self.unsuccessful_teryts)

        successful_sends_summary = "POMYŚLNE WYSŁANIA:\n"
        for authority_teryt, authority_mail in self.successful_teryts.items():
            successful_sends_summary += f"Dla gminy o TERYT {authority_teryt} na adres(y): {authority_mail}\n"

        unsuccessful_sends_summary = "WYSŁANIA ZAKOŃCZONE NIEPOWODZENIEM:\n"
        for authority_teryt, operation_summary in self.unsuccessful_teryts.items():
            unsuccessful_sends_summary += (f"Dla gminy o TERYT {authority_teryt} na adres(y): {operation_summary[0]}. "
                                         f"Powód: {operation_summary[1]}\n")

        return (f"Hejo,\nPomyślnie wysłano {successful_sends_len} mail'i z wnioskami. Nie udało się wysłać "
                f"{unsuccessful_sends_len} mail'i.\nPodsumowanie:\n\n{successful_sends_summary}\n\n"
                f"{unsuccessful_sends_summary}\n\n\nPozdrawiam serdecznie, miłego dnia i smacznej kawusi,\n"
                f"mlAI aka Sebastian Bartuzi")

    def send_confirmation(self, outlook: CDispatch):
        try:

            mail = outlook.CreateItem(0)

            mail.To = self.confirmation_addressees
            mail.Subject = self._get_confirmation_mail_title()
            mail.Body = self._get_confirmation_mail_content()

            mail.SentOnBehalfOfName = self.sender_email

            mail.Send()
            print(f"E-mail potwierdzający do adresatów {self.confirmation_addressees} wysłano pomyślnie.")

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a potwierdzającego do adresatów {self.confirmation_addressees}: {e}")

