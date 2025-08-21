from typing import List, Dict, Any

from src.models import ResponseMailModel


class ResponseService:

    def extract_mail_analysis_response_data(self, response_data: Dict[str, Any], mail_data: ResponseMailModel) -> None:

        mail_data.set_teryt(response_data.get("teryt", ""))

        response_type = response_data.get("response_type", {})
        questions_responses = response_data.get("questions_responses", {})

        mail_data.set_automatic_response(response_type.get("automatic_response", None))
        mail_data.set_mail_not_delivered(response_type.get("mail_not_delivered", None))
        mail_data.set_wrong_addressee(response_type.get("wrong_addressee", None))
        mail_data.set_action_required(response_type.get("action_required", None))
        mail_data.set_deadline_extended(response_type.get("deadline_extended", None))
        mail_data.set_refused_to_answer_fully(response_type.get("refused_to_answer_fully", None))
        mail_data.set_refused_to_answer_partially(response_type.get("refused_to_answer_partially", None))
        mail_data.set_part_answered_separately(response_type.get("part_answered_separately", None))
        mail_data.set_no_information(response_type is {} and questions_responses is {})
        mail_data.set_other_error(response_type.get("other_error", None))
        mail_data.set_additional_info_response_type(response_type.get("additional_info", None))

        mail_data.set_offers_internships(questions_responses.get("offers_internships", None))
        mail_data.set_are_internships_paid(questions_responses.get("are_internships_paid", None))
        mail_data.set_plans_paid_internships(questions_responses.get("plans_paid_internships", None))
        mail_data.set_paid_internships_number(questions_responses.get("paid_internships_number", None))
        mail_data.set_internships_salaries(questions_responses.get("internships_salaries", None))
        mail_data.set_additional_info_questions_responses(questions_responses.get("additional_info", None))

    def set_matched_teryt(self, response_data: List[str], mail_data: ResponseMailModel) -> None:

        mail_data.set_teryt(','.join(response_data))
        mail_data.set_no_teryt_matched(len(response_data) == 0)
        mail_data.set_multiple_teryts_matched(len(response_data) > 1)
