from typing import List

from win32com.client import CDispatch

from src.core import config
from src.enums import Flag
from src.models import AuthorityModel, ResponseMailModel
from src.utils import Utils, FileHandler, PathCreator


class MailService:
    def __init__(self) -> None:
        self._file_handler = FileHandler()
        self._path_creator = PathCreator()
        self._utils = Utils()
        self._successful_send_teryts = {}
        self._unsuccessful_send_teryts = {}

    def _get_authority_mail_title(self, authority_data: AuthorityModel) -> str:
        return f'Wniosek o udostępnienie informacji publicznej ({authority_data.get_authority_teryt()})'

    def _get_authority_mail_content(self, authority_data: AuthorityModel) -> str:
        salutation = self._utils.get_salutation_vocative(
            authority_data.get_authority_teryt(), authority_data.get_governor_gender()
        )
        return (f"{salutation},\nW załączniku przesyłam wniosek o udostępnienie informacji publicznej.\nW tytule "
                f"odpowiedzi zwrotnej proszę o zawarcie kodu TERYT gminy: {authority_data.get_authority_teryt()}.\n\n"
                f"Z poważaniem\nSebastian Bartuzi\nCzłonek Stowarzyszenia Młoda Lewica RP")

    def issue_authority_email(self, authority_data: AuthorityModel, outlook: CDispatch, test_mode: bool) -> None:

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
            self._successful_send_teryts.update({authority_data.get_authority_teryt(): mail_to})

            self._file_handler.remove_file(attachment_path)

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a do gminy o TERYT {authority_data.get_authority_teryt()}: {e}")
            self._unsuccessful_send_teryts.update({authority_data.get_authority_teryt(): (mail_to, e)})

    def _get_send_confirmation_mail_title(self) -> str:
        return f"mlAI: Potwierdzenie wysłania mail'i z dnia {self._utils.get_today_date()}"

    def _get_send_confirmation_mail_content(self) -> str:
        successful_sends_len = len(self._successful_send_teryts)
        unsuccessful_sends_len = len(self._unsuccessful_send_teryts)

        successful_sends_summary = "POMYŚLNE WYSŁANIA:\n"
        for authority_teryt, authority_mail in self._successful_send_teryts.items():
            successful_sends_summary += f"Dla gminy o TERYT {authority_teryt} na adres(y): {authority_mail}\n"

        unsuccessful_sends_summary = "WYSŁANIA ZAKOŃCZONE NIEPOWODZENIEM:\n"
        for authority_teryt, operation_summary in self._unsuccessful_send_teryts.items():
            unsuccessful_sends_summary += (f"Dla gminy o TERYT {authority_teryt} na adres(y): {operation_summary[0]}. "
                                         f"Powód: {operation_summary[1]}\n")

        return (f"Hejo,\n"
                f"Pomyślnie wysłano {successful_sends_len} mail'i z wnioskami.\n"
                f"Nie udało się wysłać {unsuccessful_sends_len} mail'i.\n"
                f"Podsumowanie:\n\n"
                f"{unsuccessful_sends_summary}\n\n"
                f"{successful_sends_summary}\n\n\n"
                f"Pozdrawiam serdecznie, miłego dnia i smacznej kawusi,\n"
                f"mlAI aka Sebastian Bartuzi")

    def issue_send_confirmation(self, outlook: CDispatch):

        try:

            mail = outlook.CreateItem(0)

            mail.To = config.confirmation_addressees
            mail.Subject = self._get_send_confirmation_mail_title()
            mail.Body = self._get_send_confirmation_mail_content()

            mail.SentOnBehalfOfName = config.sender_email

            mail.Send()
            print(f"E-mail potwierdzający do adresatów {config.confirmation_addressees} wysłano pomyślnie.")

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a potwierdzającego do adresatów {config.confirmation_addressees}: {e}")

    def _get_forward_attention_mail_title(self, mail_data: ResponseMailModel) -> str:
        return f"TERYT: {mail_data.get_teryt()}"

    def _get_forward_attention_mail_content(self, mail_data: ResponseMailModel, original_mail: CDispatch,
                                            forward_mail: CDispatch) -> str:

        headline_content = (f"{mail_data.get_attention_information()}\n"
                            f"Dodatkowe informacje: {mail_data.get_additional_info_response_type()}\n\n"
                            f"__________\n\n\n")

        if original_mail.HTMLBody:
            headline_content_html = f"<p>{headline_content.replace('\n', '<br>')}</p>"
            forward_mail.HTMLBody = headline_content_html + original_mail.HTMLBody

        return headline_content + original_mail.Body

    def forward_attention_needed_mails(self, outlook: CDispatch, mails_data: List[ResponseMailModel]) -> None:

        try:

            mails_to_forward_data = [mail_data for mail_data in mails_data if mail_data.get_attention_needed()]

            if not mails_to_forward_data:
                print("Brak e-maili do przekazania.")

            else:

                print(f"Rozpoczynam przekazywanie {len(mails_to_forward_data)} e-maili wymagających uwagi.")

                namespace = outlook.GetNamespace("MAPI")

                for mail_data_to_forward in mails_to_forward_data:
                    original_mail_item = None

                    if mail_data_to_forward.get_mail_id():
                        original_mail_item = namespace.GetItemFromID(mail_data_to_forward.get_mail_id())

                    if original_mail_item:

                        forward_mail = original_mail_item.Forward()

                        forward_mail.To = config.confirmation_addressees
                        forward_mail.Subject = self._get_forward_attention_mail_title(mail_data_to_forward)
                        forward_mail.Body = self._get_forward_attention_mail_content(
                            mail_data_to_forward, original_mail_item, forward_mail
                        )

                        forward_mail.Send()

                        print(f"Pomyślnie przekazano e-mail: '{mail_data_to_forward.get_mail_title()}'")

        except Exception as e:
            print(f"Błąd podczas przekazywania dalej e-mail'i wymagających uwagi forward_attention_needed_mails: {e}")


    def _get_read_confirmation_mail_title(self) -> str:
        return f"mlAI: Potwierdzenie zaktualizowania raportu z dnia {self._utils.get_today_date()}"

    def _get_read_confirmation_mail_content(self, mails_data: List[ResponseMailModel]) -> str:

        ignored_mails = 0
        attention_needed_mails = 0
        attention_needed_summary = ""
        response_mails = 0
        response_summary = ""

        for mail_data in mails_data:

            if mail_data.get_automatic_response():
                ignored_mails += 1

            if mail_data.get_attention_needed():
                attention_needed_mails += 1
                attention_needed_summary += (f"    Tytuł: {mail_data.get_mail_title()}\n"
                                             f"    Data: {mail_data.get_mail_date()}\n"
                                             f"    Powód: {mail_data.get_attention_information()}\n"
                                             f"    Dod. info: {mail_data.get_additional_info_response_type()}\n\n")

            if mail_data.get_answers_given():
                response_mails += 1
                response_summary += (f"    Tytuł: {mail_data.get_mail_title()}\n"
                                     f"    Data: {mail_data.get_mail_date()}\n\n")

        return (f"Hejo,\n\n"
                f"Otrzymano {len(mails_data)} wiadomości od {mails_data[0].get_mail_date()}.\n"
                f"Zignorowano {ignored_mails} wiadomości (automatyczne odpowiedzi).\n\n"
                f"{attention_needed_mails} wiadomości wymaga Twojej uwagi (wszystkie do Ciebie przekierowałem):\n"
                f"{attention_needed_summary}"
                f"{response_mails} zawierają odpowiedzi od gmin:\n"
                f"{response_summary}\n\n\n"
                f"Pozdrawiam serdecznie, miłego dnia i smacznej kawusi,\n"
                f"mlAI aka Sebastian Bartuzi")

    def issue_read_confirmation(self, outlook: CDispatch, mails_data: List[ResponseMailModel]) -> None:

        try:

            mail = outlook.CreateItem(0)

            mail.To = config.confirmation_addressees
            mail.Subject = self._get_read_confirmation_mail_title()
            mail.Body = self._get_read_confirmation_mail_content(mails_data)

            mail.Attachments.Add(self._path_creator.get_report_excel_file_path())

            mail.SentOnBehalfOfName = config.sender_email

            mail.Send()

            print(f"E-mail potwierdzający do adresatów {config.confirmation_addressees} wysłano pomyślnie.")

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a potwierdzającego do adresatów {config.confirmation_addressees}: {e}")

    def issue_stop_flag(self, outlook: CDispatch) -> None:

        try:

            mail = outlook.CreateItem(0)

            mail.To = config.sender_email
            mail.Subject = Flag.STOP.value
            mail.Body = Flag.STOP.value

            mail.SentOnBehalfOfName = config.sender_email

            mail.Send()

            print("E-mail z flagą STOP wysłano pomyślnie.")

        except Exception as e:
            print(f"Nie udało się wysłać e-mail'a z flagą STOP do {config.sender_email}: {e}")

    def _is_mail(self, item: CDispatch) -> bool:
        return item.Class == 43

    def read_new_emails_from_inbox(self, outlook: CDispatch) -> List[ResponseMailModel]:

        emails_data = []

        try:

            namespace = outlook.GetNamespace("MAPI")
            inbox_folder = namespace.GetDefaultFolder(6)
            items = list(inbox_folder.Items)
            items.reverse()

            for mail_ind, item in enumerate(items[6:105]):

                if self._is_mail(item):

                    mail_title = item.Subject

                    if mail_title == Flag.STOP.value:
                        break

                    mail_data = ResponseMailModel()

                    mail_data.set_mail_id(item.EntryID)
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

                                attachment_path = self._path_creator.get_attachment_path(
                                    mail_ind, attachment_id, self._file_handler.get_file_extension(attachment_file_name)
                                )

                                attachment.SaveAsFile(attachment_path)

                                attachment_path = self._file_handler.convert_docx_to_txt(attachment_path)
                                attachments_paths.append(attachment_path)

                                print(f"    Zapisano załącznik {attachment_file_name} jako {attachment_path}.")

                            else:

                                unacceptable_attachments_names.append(attachment_file_name)

                    mail_data.set_attachments_paths(attachments_paths)
                    mail_data.set_unacceptable_attachments_names(unacceptable_attachments_names)

                    print(f"Pomyślnie przeczytano e-mail nr {mail_ind}!")

                    emails_data.append(mail_data)

            print(f"Pomyślnie przeczytano {len(emails_data)} e-maili.")

        except Exception as e:
            print(f"Błąd podczas czytania e-maili: {e}")

        emails_data.reverse()
        return emails_data

