from typing import List, Optional

from src.enums.attention_information import AttentionInformation
from src.utils import FileHandler


class ResponseMailModel:
    
    _mail_id: Optional[str] = None

    _mail_sender: str = ""
    _mail_title: str = ""
    _mail_content: str = ""
    _mail_date: str = ""
    _attachments_paths: List[str] = []
    _unacceptable_attachments_names: List[str] = []

    _teryt: str = ""
    _no_teryt_matched: bool = False
    _multiple_teryts_matched: bool = False

    _automatic_response: Optional[bool] = None
    _mail_not_delivered: Optional[bool] = None
    _wrong_addressee: Optional[bool] = None
    _action_required: Optional[bool] = None
    _deadline_extended: Optional[str] = None
    _refused_to_answer_fully: Optional[bool] = None
    _refused_to_answer_partially: Optional[bool] = None
    _part_answered_separately: Optional[bool] = None
    _additional_info_response_type: Optional[str] = None

    _offers_internships: Optional[bool] = None
    _are_internships_paid: Optional[bool] = None
    _plans_paid_internships: Optional[bool] = None
    _paid_internships_number: Optional[float] = None
    _internships_salaries: Optional[float] = None
    _additional_info_questions_responses: Optional[str] = None

    _attention_information: Optional[str] = None


    def __init__(self):
        self.file_handler = FileHandler()
        
    def set_mail_id(self, mail_id: str):

        if not isinstance(mail_id, str):
            raise ValueError(f"Nieprawidłowa wartość mail_id: {mail_id}!")

        self._mail_id = mail_id

    def get_mail_id(self) -> str:
        return self._mail_id

    def set_mail_title(self, mail_title: str) -> None:

        if not isinstance(mail_title, str):
            raise ValueError(f"Nieprawidłowa wartość mail_title: {mail_title}!")

        self._mail_title = mail_title

    def get_mail_title(self) -> str:
        return self._mail_title

    def set_mail_content(self, mail_content: str) -> None:

        if not isinstance(mail_content, str):
            raise ValueError(f"Nieprawidłowa wartość mail_content: {mail_content} dla wiadomości {self._mail_id}!")

        self._mail_content = mail_content

    def get_mail_content(self) -> str:
        return self._mail_content
    
    def set_mail_sender(self, mail_sender: str) -> None:

        if not isinstance(mail_sender, str) or not mail_sender:
            raise ValueError(f"Nieprawidłowa wartość mail_sender: {mail_sender} dla wiadomości {self._mail_id}!")

        self._mail_sender = mail_sender

    def get_mail_sender(self) -> str:
        return self._mail_sender

    def set_mail_date(self, mail_date: str) -> None:

        if not isinstance(mail_date, str) or not mail_date:
            raise ValueError(f"Nieprawidłowa wartość mail_date: {mail_date} dla wiadomości {self._mail_id}!")

        self._mail_date = mail_date

    def get_mail_date(self) -> str:
        return self._mail_date

    def set_attachments_paths(self, attachments_paths: List[str]) -> None:

        if not isinstance(attachments_paths, List):
            raise ValueError(f"Wartość attachments_paths dla wiadomości {self._mail_id} musi być listą!")

        for attachment_path in attachments_paths:
            if not isinstance(attachment_path, str) or not self.file_handler.check_file_exists(attachment_path):
                raise ValueError(f"Dla wiadomości {self._mail_id} nieprawidłowa ścieżka załącznika lub plik nie "
                                 f"istnieje: {attachment_path}!")

        self._attachments_paths = attachments_paths

    def get_attachments_paths(self) -> List[str]:
        return self._attachments_paths

    def set_unacceptable_attachments_names(self, unacceptable_attachments_names: List[str]) -> None:

        if not isinstance(unacceptable_attachments_names, List):
            raise ValueError(f"Wartość unacceptable_attachments_names dla wiadomości {self._mail_id} musi być listą!")

        for unacceptable_attachment_name in unacceptable_attachments_names:
            if not isinstance(unacceptable_attachment_name, str):
                raise ValueError(f"Dla wiadomości {self._mail_id} nieprawidłowa ścieżka załącznika lub plik nie "
                                 f"istnieje: {unacceptable_attachment_name}!")

        self._unacceptable_attachments_names = unacceptable_attachments_names

    def get_unacceptable_attachments_names(self) -> List[str]:
        return self._unacceptable_attachments_names
    
    def set_teryt(self, teryt: str) -> None:

        if not isinstance(teryt, str):
            raise ValueError(f"Nieprawidłowa wartość teryt: {teryt} dla wiadomości {self._mail_id}!")

        self._teryt = teryt

    def get_teryt(self) -> str:
        return self._teryt

    def set_no_teryt_matched(self, no_teryt_matched: bool) -> None:

        if not isinstance(no_teryt_matched, bool):
            raise ValueError(f"Nieprawidłowa wartość no_teryt_matched: {no_teryt_matched} dla wiadomości "
                             f"{self._mail_id}!")

        self._no_teryt_matched = no_teryt_matched

    def get_no_teryt_matched(self) -> bool:
        return self._no_teryt_matched

    def set_multiple_teryts_matched(self, multiple_teryts_matched: bool) -> None:

        if not isinstance(multiple_teryts_matched, bool):
            raise ValueError(f"Nieprawidłowa wartość multiple_teryts_matched: {multiple_teryts_matched} dla wiadomości "
                             f"{self._mail_id}!")

        self._multiple_teryts_matched = multiple_teryts_matched

    def get_multiple_teryts_matched(self) -> bool:
        return self._multiple_teryts_matched

    def set_automatic_response(self, automatic_response: Optional[bool]) -> None:

        if not isinstance(automatic_response, bool) and automatic_response is not None:
            raise ValueError(f"Nieprawidłowa wartość automatic_response: {automatic_response} dla wiadomości "
                             f"{self._mail_id}!")

        self._automatic_response = automatic_response

    def get_automatic_response(self) -> Optional[bool]:
        return self._automatic_response

    def set_mail_not_delivered(self, mail_not_delivered: Optional[bool]) -> None:

        if not isinstance(mail_not_delivered, bool) and mail_not_delivered is not None:
            raise ValueError(f"Nieprawidłowa wartość mail_not_delivered: {mail_not_delivered} dla wiadomości "
                             f"{self._mail_id}!")

        self._mail_not_delivered = mail_not_delivered

    def get_mail_not_delivered(self) -> Optional[bool]:
        return self._mail_not_delivered

    def set_wrong_addressee(self, wrong_addressee: Optional[bool]) -> None:

        if not isinstance(wrong_addressee, bool) and wrong_addressee is not None:
            raise ValueError(f"Nieprawidłowa wartość wrong_addressee: {wrong_addressee} dla wiadomości "
                             f"{self._mail_id}!")

        self._wrong_addressee = wrong_addressee

    def get_wrong_addressee(self) -> Optional[bool]:
        return self._wrong_addressee

    def set_action_required(self, action_required: Optional[bool]) -> None:

        if not isinstance(action_required, bool) and action_required is not None:
            raise ValueError(f"Nieprawidłowa wartość action_required: {action_required} dla wiadomości "
                             f"{self._mail_id}!")

        self._action_required = action_required

    def get_action_required(self) -> Optional[bool]:
        return self._action_required

    def set_deadline_extended(self, deadline_extended: Optional[str]) -> None:

        if not isinstance(deadline_extended, str) and deadline_extended is not None:
            raise ValueError(f"Nieprawidłowa wartość deadline_extended: {deadline_extended} dla wiadomości "
                             f"{self._mail_id}!")

        self._deadline_extended = deadline_extended

    def get_deadline_extended(self) -> Optional[str]:
        return self._deadline_extended

    def set_refused_to_answer_fully(self, refused_to_answer_fully: Optional[bool]) -> None:

        if not isinstance(refused_to_answer_fully, bool) and refused_to_answer_fully is not None:
            raise ValueError(f"Nieprawidłowa wartość refused_to_answer_fully: {refused_to_answer_fully} dla wiadomości "
                             f"{self._mail_id}!")

        self._refused_to_answer_fully = refused_to_answer_fully

    def get_refused_to_answer_fully(self) -> Optional[bool]:
        return self._refused_to_answer_fully

    def set_refused_to_answer_partially(self, refused_to_answer_partially: Optional[bool]) -> None:

        if not isinstance(refused_to_answer_partially, bool) and refused_to_answer_partially is not None:
            raise ValueError(f"Nieprawidłowa wartość refused_to_answer_partially: {refused_to_answer_partially} dla "
                             f"wiadomości {self._mail_id}!")

        self._refused_to_answer_partially = refused_to_answer_partially

    def get_refused_to_answer_partially(self) -> Optional[bool]:
        return self._refused_to_answer_partially

    def set_part_answered_separately(self, part_answered_separately: Optional[bool]) -> None:

        if not isinstance(part_answered_separately, bool) and part_answered_separately is not None:
            raise ValueError(f"Nieprawidłowa wartość part_answered_separately: {part_answered_separately} dla "
                             f"wiadomości {self._mail_id}!")

        self._part_answered_separately = part_answered_separately

    def get_part_answered_separately(self) -> Optional[bool]:
        return self._part_answered_separately

    def set_additional_info_response_type(self, additional_info_response_type: Optional[str]) -> None:

        if not isinstance(additional_info_response_type, str) and additional_info_response_type is not None:
            raise ValueError(f"Nieprawidłowa wartość additional_info_response_type: {additional_info_response_type} "
                             f"dla wiadomości {self._mail_id}!")

        self._additional_info_response_type = additional_info_response_type

    def get_additional_info_response_type(self) -> Optional[str]:
        return self._additional_info_response_type

    def set_offers_internships(self, offers_internships: Optional[bool]) -> None:

        if not isinstance(offers_internships, bool) and offers_internships is not None:
            raise ValueError(f"Nieprawidłowa wartość offers_internships: {offers_internships} dla wiadomości "
                             f"{self._mail_id}!")

        self._offers_internships = offers_internships

    def get_offers_internships(self) -> Optional[bool]:
        return self._offers_internships

    def set_are_internships_paid(self, are_internships_paid: Optional[bool]) -> None:

        if not isinstance(are_internships_paid, bool) and are_internships_paid is not None:
            raise ValueError(f"Nieprawidłowa wartość are_internships_paid: {are_internships_paid} dla wiadomości "
                             f"{self._mail_id}!")

        self._are_internships_paid = are_internships_paid

    def get_are_internships_paid(self) -> Optional[bool]:
        return self._are_internships_paid

    def set_plans_paid_internships(self, plans_paid_internships: Optional[bool]) -> None:

        if not isinstance(plans_paid_internships, bool) and plans_paid_internships is not None:
            raise ValueError(f"Nieprawidłowa wartość plans_paid_internships: {plans_paid_internships} dla wiadomości "
                             f"{self._mail_id}!")

        self._plans_paid_internships = plans_paid_internships

    def get_plans_paid_internships(self) -> Optional[bool]:
        return self._plans_paid_internships

    def set_paid_internships_number(self, paid_internships_number: Optional[float]) -> None:

        if not isinstance(paid_internships_number, float) and paid_internships_number is not None:
            raise ValueError(f"Nieprawidłowa wartość paid_internships_number: {paid_internships_number} dla wiadomości "
                             f"{self._mail_id}!")

        self._paid_internships_number = paid_internships_number

    def get_paid_internships_number(self) -> Optional[float]:
        return self._paid_internships_number

    def set_internships_salaries(self, internships_salaries: Optional[float]) -> None:

        if not isinstance(internships_salaries, float) and internships_salaries is not None:
            raise ValueError(f"Nieprawidłowa wartość internships_salaries: {internships_salaries} dla wiadomości "
                             f"{self._mail_id}!")

        self._internships_salaries = internships_salaries

    def get_internships_salaries(self) -> Optional[float]:
        return self._internships_salaries

    def set_additional_info_questions_responses(self, additional_info_questions_responses: Optional[str]) -> None:

        if not isinstance(additional_info_questions_responses, str) and additional_info_questions_responses is not None:
            raise ValueError(f"Nieprawidłowa wartość additional_info_questions_responses: "
                             f"{additional_info_questions_responses} dla wiadomości {self._mail_id}!")

        self._additional_info_questions_responses = additional_info_questions_responses

    def get_additional_info_questions_responses(self) -> Optional[str]:
        return self._additional_info_questions_responses

    def get_attention_information(self) -> Optional[str]:
        return (
            AttentionInformation.NO_TERYT_MATCHED.value if self._no_teryt_matched else
            AttentionInformation.MULTIPLE_TERYTS_MATCHED.value if self._multiple_teryts_matched else
            AttentionInformation.MAIL_NOT_DELIVERED.value if self._mail_not_delivered else
            AttentionInformation.WRONG_ADDRESSEE.value if self._wrong_addressee else
            AttentionInformation.ACTION_REQUIRED.value if self._action_required else
            AttentionInformation.DEADLINE_EXTENDED.value if self._deadline_extended else
            AttentionInformation.REFUSED_TO_ANSWER_FULLY.value if self._refused_to_answer_fully else
            AttentionInformation.REFUSED_TO_ANSWER_PARTIALLY.value if self._refused_to_answer_partially else
            AttentionInformation.PART_ANSWERED_SEPARATELY.value if self._part_answered_separately else
            ""
        )

    def get_attention_needed(self) -> bool:
        return any([self._mail_not_delivered, self._wrong_addressee, self._action_required, self._deadline_extended,
                    self._refused_to_answer_fully, self._refused_to_answer_partially, self._part_answered_separately,
                    self._no_teryt_matched, self._multiple_teryts_matched])

    def get_answers_given(self) -> bool:
        return any(flag is not None for flag in [
            self._offers_internships, self._are_internships_paid, self._plans_paid_internships,
            self._paid_internships_number, self._internships_salaries
        ])