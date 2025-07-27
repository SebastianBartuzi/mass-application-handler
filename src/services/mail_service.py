from typing import List

from win32com.client import CDispatch

from src.core import config
from src.enums import Flag
from src.models import AuthorityModel, ResponseModel
from src.utils import Utils, FileHandler, PathCreator


class MailService:
    def __init__(self) -> None:
        self._file_handler = FileHandler()
        self._path_creator = PathCreator()
        self._utils = Utils()
        self._successful_teryts = {}
        self._unsuccessful_teryts = {}

    def _get_authority_mail_title(self, authority_data: AuthorityModel) -> str:
        return f'Wniosek o udostępnienie informacji publicznej ({authority_data.get_authority_teryt()})'

    def _get_authority_mail_content(self, authority_data: AuthorityModel) -> str:
        salutation = self._utils.get_salutation_vocative(
            authority_data.get_authority_teryt(), authority_data.get_governor_gender()
        )
        return (f"{salutation},\nW załączniku przesyłam wniosek o udostępnienie informacji publicznej.\nW tytule "
                f"odpowiedzi zwrotnej proszę o zawarcie kodu TERYT gminy: {authority_data.get_authority_teryt()}.\n\n"
                f"Z poważaniem\nSebastian Bartuzi\nCzłonek Stowarzyszenia Młoda Lewica RP")

    def send_authority_email(self, authority_data: AuthorityModel, outlook: CDispatch, test_mode: bool) -> None:

        if test_mode:
            mail_to = config.confirmation_addressees
        else:
            mail_to = authority_data.get_authority_office_mail()

        try:

            mail = outlook.CreateItem(0)

            mail.To = mail_to
            mail.Subject = self._get_authority_mail_title(authority_data)
            mail.Body = self._get_authority_mail_content(authority_data)

            mail.SentOnBehalfOfName = config.sender_email

            attachment_path = authority_data.get_application_pdf_path()
            mail.Attachments.Add(attachment_path)

            mail.Send()
            print(f"E-mail do gminy o TERYT {authority_data.get_authority_teryt()} wysłano pomyślnie.")
            self._successful_teryts.update({authority_data.get_authority_teryt(): mail_to})

            self._file_handler.remove_file(attachment_path)

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a do gminy o TERYT {authority_data.get_authority_teryt()}: {e}")
            self._unsuccessful_teryts.update({authority_data.get_authority_teryt(): (mail_to, e)})

    def _get_confirmation_mail_title(self) -> str:
        return f"mlAI: Potwierdzenie operacji z dnia {self._utils.get_today_date()}"

    def _get_confirmation_mail_content(self) -> str:
        successful_sends_len = len(self._successful_teryts)
        unsuccessful_sends_len = len(self._unsuccessful_teryts)

        successful_sends_summary = "POMYŚLNE WYSŁANIA:\n"
        for authority_teryt, authority_mail in self._successful_teryts.items():
            successful_sends_summary += f"Dla gminy o TERYT {authority_teryt} na adres(y): {authority_mail}\n"

        unsuccessful_sends_summary = "WYSŁANIA ZAKOŃCZONE NIEPOWODZENIEM:\n"
        for authority_teryt, operation_summary in self._unsuccessful_teryts.items():
            unsuccessful_sends_summary += (f"Dla gminy o TERYT {authority_teryt} na adres(y): {operation_summary[0]}. "
                                         f"Powód: {operation_summary[1]}\n")

        return (f"Hejo,\nPomyślnie wysłano {successful_sends_len} mail'i z wnioskami. Nie udało się wysłać "
                f"{unsuccessful_sends_len} mail'i.\nPodsumowanie:\n\n{successful_sends_summary}\n\n"
                f"{unsuccessful_sends_summary}\n\n\nPozdrawiam serdecznie, miłego dnia i smacznej kawusi,\n"
                f"mlAI aka Sebastian Bartuzi")

    def send_confirmation(self, outlook: CDispatch):
        try:

            mail = outlook.CreateItem(0)

            mail.To = config.confirmation_addressees
            mail.Subject = self._get_confirmation_mail_title()
            mail.Body = self._get_confirmation_mail_content()

            mail.SentOnBehalfOfName = config.sender_email

            mail.Send()
            print(f"E-mail potwierdzający do adresatów {config.confirmation_addressees} wysłano pomyślnie.")

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a potwierdzającego do adresatów {config.confirmation_addressees}: {e}")

    def _is_mail(self, item: CDispatch) -> bool:
        return item.Class == 43

    def read_emails_from_inbox(self, outlook: CDispatch) -> List[ResponseModel]:

        emails_data = []

        try:

            namespace = outlook.GetNamespace("MAPI")
            folder = namespace.GetDefaultFolder(6)

            for mail_id, item in enumerate(reversed(folder.Items)):

                if self._is_mail(item):

                    mail_title = item.Subject

                    if mail_title == Flag.STOP.value:
                        break

                    mail_data = ResponseModel()

                    mail_data.set_mail_title(mail_title)
                    mail_data.set_mail_content(item.Body)
                    mail_data.set_mail_sender(item.SenderEmailAddress)
                    mail_data.set_mail_date(item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"))

                    print(f"Czytam e-mail: From='{mail_data.get_mail_sender()}',"
                          f"Subject='{mail_data.get_mail_title()}', Received='{mail_data.get_mail_date()}'")

                    attachments_paths = []
                    unacceptable_attachments_names = []

                    if item.Attachments.Count > 0:

                        for attachment_id, attachment in enumerate(item.Attachments):

                            attachment_file_name = attachment.FileName

                            if self._file_handler.is_gemini_accepted_file(attachment_file_name):

                                print(attachment.FileName)

                                attachment_path = self._path_creator.get_attachment_path(
                                    mail_id, attachment_id, self._file_handler.get_file_extension(attachment_file_name)
                                )

                                attachment.SaveAsFile(attachment_path)
                                attachments_paths.append(attachment_path)
                                print(f"    Zapisano załącznik: {attachment_file_name}")

                            else:

                                unacceptable_attachments_names.append(attachment_file_name)

                    mail_data.set_attachments_paths(attachments_paths)
                    mail_data.set_unacceptable_attachments_names(unacceptable_attachments_names)

                    print("Pomyślnie przeczytano e-mail!")

            print(f"Pomyślnie przeczytano {len(emails_data)} e-maili.")

        except Exception as e:
            print(f"Błąd podczas czytania e-maili: {e}")

        return emails_data

