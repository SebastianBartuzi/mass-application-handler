from typing import List

from src.utils import FileHandler


class ResponseModel:

    _mail_sender: str = ""
    _mail_title: str = ""
    _mail_content: str = ""
    _mail_date: str = ""
    _attachments_paths: List[str] = []
    _unacceptable_attachments_names: List[str] = []
    _teryt: str = ""
    _automatic_response: bool = None
    _mail_not_delivered: bool = None
    _action_required: bool = None
    _deadline_extended: str = None
    _refused_to_answer_fully: bool = None
    _refused_to_answer_partially: bool = None
    _part_answered_separately: bool = None
    _additional_info_response_type: str = None
    _offers_internships: bool = None
    _are_internships_paid: bool = None
    _plans_paid_internships: bool = None
    _internships_number: float = None
    _internships_salaries: float = None
    _additional_info_response_data: str = None


    def __init__(self):
        self.file_handler = FileHandler()

    def set_mail_title(self, mail_title: str) -> None:

        if not isinstance(mail_title, str):
            raise ValueError(f"Nieprawidłowa wartość mail_title: {mail_title}!")

        self._mail_title = mail_title

    def get_mail_title(self) -> str:
        return self._mail_title

    def set_mail_content(self, mail_content: str) -> None:

        if not isinstance(mail_content, str):
            raise ValueError(f"Nieprawidłowa wartość mail_content: {mail_content} dla wiadomości {self._mail_title}!")

        self._mail_content = mail_content

    def get_mail_content(self) -> str:
        return self._mail_content
    
    def set_mail_sender(self, mail_sender: str) -> None:

        if not isinstance(mail_sender, str) or not mail_sender:
            raise ValueError(f"Nieprawidłowa wartość mail_sender: {mail_sender} dla wiadomości {self._mail_title}!")

        self._mail_sender = mail_sender

    def get_mail_sender(self) -> str:
        return self._mail_sender

    def set_mail_date(self, mail_date: str) -> None:

        if not isinstance(mail_date, str) or not mail_date:
            raise ValueError(f"Nieprawidłowa wartość mail_date: {mail_date} dla wiadomości {self._mail_title}!")

        self._mail_date = mail_date

    def get_mail_date(self) -> str:
        return self._mail_date

    def set_attachments_paths(self, attachments_paths: List[str]) -> None:

        if not isinstance(attachments_paths, List):
            raise ValueError(f"Wartość attachments_paths dla wiadomości {self._mail_title} musi być listą!")

        for attachment_path in attachments_paths:
            if not isinstance(attachment_path, str) or not self.file_handler.check_file_exists(attachment_path):
                raise ValueError(f"Dla wiadomości {self._mail_title} nieprawidłowa ścieżka załącznika lub plik nie "
                                 f"istnieje: {attachment_path}!")

        self._attachments_paths = attachments_paths

    def get_attachments_paths(self) -> List[str]:
        return self._attachments_paths

    def set_unacceptable_attachments_names(self, unacceptable_attachments_names: List[str]) -> None:

        if not isinstance(unacceptable_attachments_names, List):
            raise ValueError(f"Wartość unacceptable_attachments_names dla wiadomości {self._mail_title} musi być "
                             f"listą!")

        for unacceptable_attachment_name in unacceptable_attachments_names:
            if not isinstance(unacceptable_attachment_name, str):
                raise ValueError(f"Dla wiadomości {self._mail_title} nieprawidłowa ścieżka załącznika lub plik nie "
                                 f"istnieje: {unacceptable_attachment_name}!")

        self._unacceptable_attachments_names = unacceptable_attachments_names

    def get_unacceptable_attachments_names(self) -> List[str]:
        return self._unacceptable_attachments_names
    
    def set_teryt(self, teryt: str) -> None:

        if not isinstance(teryt, str):
            raise ValueError(f"Nieprawidłowa wartość teryt: {teryt} dla wiadomości {self._mail_title}!")

        self._teryt = teryt

    def get_teryt(self) -> str:
        return self._teryt

    def set_automatic_response(self, automatic_response: bool) -> None:

        if not isinstance(automatic_response, bool):
            raise ValueError(f"Nieprawidłowa wartość automatic_response: {automatic_response} dla wiadomości "
                             f"{self._mail_title}!")

        self._automatic_response = automatic_response

    def get_automatic_response(self) -> bool:
        return self._automatic_response

    def set_mail_not_delivered(self, mail_not_delivered: bool) -> None:

        if not isinstance(mail_not_delivered, bool):
            raise ValueError(f"Nieprawidłowa wartość mail_not_delivered: {mail_not_delivered} dla wiadomości "
                             f"{self._mail_title}!")

        self._mail_not_delivered = mail_not_delivered

    def get_mail_not_delivered(self) -> bool:
        return self._mail_not_delivered

    def set_action_required(self, action_required: bool) -> None:

        if not isinstance(action_required, bool):
            raise ValueError(f"Nieprawidłowa wartość action_required: {action_required} dla wiadomości "
                             f"{self._mail_title}!")

        self._action_required = action_required

    def get_action_required(self) -> bool:
        return self._action_required

    def set_deadline_extended(self, deadline_extended: str) -> None:

        if not isinstance(deadline_extended, str):
            raise ValueError(f"Nieprawidłowa wartość deadline_extended: {deadline_extended} dla wiadomości "
                             f"{self._mail_title}!")

        self._deadline_extended = deadline_extended

    def get_deadline_extended(self) -> str:
        return self._deadline_extended

    def set_refused_to_answer_fully(self, refused_to_answer_fully: bool) -> None:

        if not isinstance(refused_to_answer_fully, bool):
            raise ValueError(f"Nieprawidłowa wartość refused_to_answer_fully: {refused_to_answer_fully} dla wiadomości "
                             f"{self._mail_title}!")

        self._refused_to_answer_fully = refused_to_answer_fully

    def get_refused_to_answer_fully(self) -> bool:
        return self._refused_to_answer_fully

    def set_refused_to_answer_partially(self, refused_to_answer_partially: bool) -> None:

        if not isinstance(refused_to_answer_partially, bool):
            raise ValueError(f"Nieprawidłowa wartość refused_to_answer_partially: {refused_to_answer_partially} dla "
                             f"wiadomości {self._mail_title}!")

        self._refused_to_answer_partially = refused_to_answer_partially

    def get_refused_to_answer_partially(self) -> bool:
        return self._refused_to_answer_partially

    def set_part_answered_separately(self, part_answered_separately: bool) -> None:

        if not isinstance(part_answered_separately, bool):
            raise ValueError(f"Nieprawidłowa wartość part_answered_separately: {part_answered_separately} dla "
                             f"wiadomości {self._mail_title}!")

        self._part_answered_separately = part_answered_separately

    def get_part_answered_separately(self) -> bool:
        return self._part_answered_separately

    def set_additional_info_response_type(self, additional_info_response_type: str) -> None:

        if not isinstance(additional_info_response_type, str):
            raise ValueError(f"Nieprawidłowa wartość additional_info_response_type: {additional_info_response_type} "
                             f"dla wiadomości {self._mail_title}!")

        self._additional_info_response_type = additional_info_response_type

    def get_additional_info_response_type(self) -> str:
        return self._additional_info_response_type

    def set_offers_internships(self, offers_internships: bool) -> None:

        if not isinstance(offers_internships, bool):
            raise ValueError(f"Nieprawidłowa wartość offers_internships: {offers_internships} dla wiadomości "
                             f"{self._mail_title}!")

        self._offers_internships = offers_internships

    def get_offers_internships(self) -> bool:
        return self._offers_internships

    def set_are_internships_paid(self, are_internships_paid: bool) -> None:

        if not isinstance(are_internships_paid, bool):
            raise ValueError(f"Nieprawidłowa wartość are_internships_paid: {are_internships_paid} dla wiadomości "
                             f"{self._mail_title}!")

        self._are_internships_paid = are_internships_paid

    def get_are_internships_paid(self) -> bool:
        return self._are_internships_paid

    def set_plans_paid_internships(self, plans_paid_internships: bool) -> None:

        if not isinstance(plans_paid_internships, bool):
            raise ValueError(f"Nieprawidłowa wartość plans_paid_internships: {plans_paid_internships} dla wiadomości "
                             f"{self._mail_title}!")

        self._plans_paid_internships = plans_paid_internships

    def get_plans_paid_internships(self) -> bool:
        return self._plans_paid_internships

    def set_internships_number(self, internships_number: float) -> None:

        if not isinstance(internships_number, float):
            raise ValueError(f"Nieprawidłowa wartość internships_number: {internships_number} dla wiadomości "
                             f"{self._mail_title}!")

        self._internships_number = internships_number

    def get_internships_number(self) -> float:
        return self._internships_number

    def set_internships_salaries(self, internships_salaries: float) -> None:

        if not isinstance(internships_salaries, float):
            raise ValueError(f"Nieprawidłowa wartość internships_salaries: {internships_salaries} dla wiadomości "
                             f"{self._mail_title}!")

        self._internships_salaries = internships_salaries

    def get_internships_salaries(self) -> float:
        return self._internships_salaries

    def set_additional_info_response_data(self, additional_info_response_data: str) -> None:

        if not isinstance(additional_info_response_data, str):
            raise ValueError(f"Nieprawidłowa wartość additional_info_response_data: {additional_info_response_data} "
                             f"dla wiadomości {self._mail_title}!")

        self._additional_info_response_data = additional_info_response_data

    def get_additional_info_response_data(self) -> str:
        return self._additional_info_response_data