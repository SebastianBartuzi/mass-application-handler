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

        # for mail_data in mails_data:
        #
        #     prompt = self._prompt_service.get_mail_analysis_prompt(mail_data)
        #     response_data = self._gemini_service.analyse_mail(prompt, mail_data)
        #
        #     self._response_service.extract_mail_analysis_response_data(response_data, mail_data)
        #
        #     if not mail_data.get_teryt() and not mail_data.get_automatic_response():
        #
        #         prompt = self._prompt_service.get_teryt_matcher_prompt(mail_data, authorities_data)
        #         response_data = self._gemini_service.match_teryt(prompt, mail_data)
        #
        #         self._response_service.set_matched_teryt(response_data, mail_data)
        #
        # self._excel_service.fill_report_excel(mails_data)
        #
        # self._mail_service.forward_attention_needed_mails(outlook, mails_data)
        # self._mail_service.issue_read_confirmation(outlook, mails_data)
        # self._mail_service.issue_stop_flag(outlook)

        responses = [
            {
                "teryt": "020101",
                "response_type": {
                    "automatic_response": True,
                    "additional_info": "Automatyczna odpowiedź o otrzymaniu wniosku."
                },
                "questions_responses": {}
            },
            {
                "teryt": "141201",
                "response_type": {
                    "deadline_extended": "2024-08-30",
                    "additional_info": "Przedłużenie terminu odpowiedzi ze względu na skomplikowany charakter sprawy."
                },
                "questions_responses": {}
            },
            {
                "teryt": "020103",
                "response_type": {
                    "refused_to_answer_fully": True,
                    "additional_info": "Odmowa udzielenia odpowiedzi ze względu na ochronę danych osobowych pracowników, którzy mieliby przeprowadzać staże."
                },
                "questions_responses": {}
            },
            {
                "teryt": "060201",
                "response_type": {},
                "questions_responses": {
                    "offers_internships": True,
                    "are_internships_paid": False,
                    "plans_paid_internships": True,
                    "paid_internships_number": 0.0,
                    "internships_salaries": 0.0,
                    "additional_info": "Urząd nie oferuje płatnych staży, ale planuje wprowadzić je w przyszłym roku budżetowym. Obecnie dostępne są tylko staże bezpłatne dla studentów trzeciego i czwartego roku."
                }
            },
            {
                "teryt": "246101",
                "response_type": {
                    "refused_to_answer_partially": True,
                    "additional_info": "Odmowa udzielenia informacji na temat liczby miejsc na staże ze względu na poufny charakter danych planistycznych. Informacje o zarobkach podano w załączniku."
                },
                "questions_responses": {
                    "offers_internships": True,
                    "are_internships_paid": True,
                    "plans_paid_internships": False,
                    "paid_internships_number": None,
                    "internships_salaries": 1500.0,
                    "additional_info": "Staże są płatne, ale liczba miejsc może się zmieniać w zależności od potrzeb poszczególnych wydziałów."
                }
            },
            {
                "teryt": "",
                "response_type": {
                    "mail_not_delivered": True,
                    "additional_info": "Adres e-mail odbiorcy jest niepoprawny lub skrzynka odbiorcza jest pełna. Wiadomość nie została dostarczona."
                },
                "questions_responses": {}
            },
            {
                "teryt": "146502",
                "response_type": {
                    "action_required": True,
                    "additional_info": "Wniosek powinien być skierowany do Biura Kadr, a nie do Sekretariatu. Proszę przesłać wniosek na poprawny adres e-mail."
                },
                "questions_responses": {}
            },
            {
                "teryt": "020101",
                "response_type": {
                    "deadline_extended": "2024-09-15",
                    "additional_info": "Wniosek wymaga dodatkowej analizy prawnej, co powoduje przedłużenie terminu odpowiedzi."
                },
                "questions_responses": {}
            },
            {
                "teryt": "120801",
                "response_type": {
                    "part_answered_separately": True,
                    "additional_info": "Część odpowiedzi dotycząca planów na przyszłość zostanie przesłana w odrębnym piśmie po konsultacji z zarządem."
                },
                "questions_responses": {
                    "offers_internships": True,
                    "are_internships_paid": True,
                    "paid_internships_number": 5.0
                }
            },
            {
                "teryt": "181403",
                "response_type": {},
                "questions_responses": {
                    "offers_internships": False,
                    "are_internships_paid": False,
                    "plans_paid_internships": False,
                    "additional_info": "Urząd nie oferuje i nie planuje oferować staży studenckich."
                }
            }
        ]

        for response_id, mail_data in enumerate(mails_data[:len(responses)]):

            response_data = responses[response_id]

            self._response_service.extract_mail_analysis_response_data(response_data, mail_data)

            if not mail_data.get_teryt() and not mail_data.get_automatic_response():

                response_data = ["040101", "060101"]

                self._response_service.set_matched_teryt(response_data, mail_data)

        self._excel_service.fill_report_excel(mails_data[:len(responses)])

        self._mail_service.forward_attention_needed_mails(outlook, mails_data[:len(responses)])
        self._mail_service.issue_read_confirmation(outlook, mails_data[:len(responses)])
        self._mail_service.issue_stop_flag(outlook)
