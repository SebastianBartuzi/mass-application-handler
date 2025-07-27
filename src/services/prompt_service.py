from typing import List

from src.core import config
from src.models import ResponseModel, AuthorityModel
from src.utils import FileHandler


class PromptService:
    def __init__(self) -> None:
        self._file_handler = FileHandler()

    def _fill_mail_placeholders(self, prompt: str, mail_data: ResponseModel) -> str:

        return prompt.replace("{{mail_from}}", mail_data.get_mail_sender()
                     .replace("{{mail_title}}", mail_data.get_mail_title())
                     .replace("{{mail_content}}", mail_data.get_mail_content())
        )

    def _fill_authorities_data(self, prompt: str, authorities_data: List[AuthorityModel]) -> str:

        authorities_data_str = ""

        for authority_data in authorities_data:

            authorities_data_str += (
                f"TERYT: {authority_data.get_authority_teryt()}, "
                f"Authority name: {authority_data.get_authority_name()}, "
                f"Authority mayor: {authority_data.get_governor_title()} {authority_data.get_governor_name()}, "
                f"Office e-mail(s) address(es): {authority_data.get_authority_office_mail()}\n"
            )

        return prompt.replace("{{authorities_data}}", authorities_data_str)

    def get_mail_analysis_prompt(self, mail_data: ResponseModel) -> str:

        prompt = self._file_handler.read_string_from_txt(config.mail_analysis_prompt_path)

        return self._fill_mail_placeholders(prompt, mail_data)

    def get_teryt_matcher_prompt(self, mail_data: ResponseModel, authorities_data: List[AuthorityModel]) -> str:

        prompt = self._file_handler.read_string_from_txt(config.teryt_matcher_prompt_path)

        prompt = self._fill_mail_placeholders(prompt, mail_data)
        return self._fill_authorities_data(prompt, authorities_data)