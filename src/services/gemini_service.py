import mimetypes

from typing import Dict, List, Any, Union

from google.ai.generativelanguage_v1 import GenerateContentResponse
from tenacity import retry, stop_after_attempt, wait_exponential, RetryCallState

from google import genai
from google.genai import types

from src.core import config
from src.models import ResponseMailModel
from src.utils import FileHandler, TypeConverter


class GeminiService:

    def __init__(self) -> None:
        self._file_handler = FileHandler()
        self._type_converter = TypeConverter()

    def _get_client(self):
        return genai.Client(
            api_key=config.gemini_api_key
        )

    def _get_contents(self, prompt: str, attachments_paths: List[str]):

        attachments = []

        for attachment_path in attachments_paths:

            mime_type, _ = mimetypes.guess_type(attachment_path)
            if mime_type is None:
                mime_type = "application/octet-stream"

            with open(attachment_path, "rb") as f:
                document_bytes = f.read()

                document_part = types.Part(
                    inline_data=types.Blob(
                        mime_type=mime_type,
                        data=document_bytes
                    )
                )
                attachments.append(document_part)

        # This is the corrected parts list construction
        # The prompt string is now also a types.Part
        all_parts = [types.Part(text=prompt), *attachments]

        return [
            types.Content(
                role="user",
                parts=all_parts
            )
        ]

    def _get_config(self, response_schema: Dict[str, Any]):
        return types.GenerateContentConfig(
            temperature=1,
            top_p=0.95,
            seed=0,
            max_output_tokens=65535,
            safety_settings=[types.SafetySetting(
                category="HARM_CATEGORY_HATE_SPEECH",
                threshold="OFF"
            ), types.SafetySetting(
                category="HARM_CATEGORY_DANGEROUS_CONTENT",
                threshold="OFF"
            ), types.SafetySetting(
                category="HARM_CATEGORY_SEXUALLY_EXPLICIT",
                threshold="OFF"
            ), types.SafetySetting(
                category="HARM_CATEGORY_HARASSMENT",
                threshold="OFF"
            )],
            response_mime_type="application/json",
            response_schema=response_schema,
            thinking_config=types.ThinkingConfig(
                thinking_budget=config.thinking_budget,
            ),
        )

    def _get_response_schema(self, file_path: str) -> Union[Dict[str, Any], List[Any]]:
        return self._file_handler.read_dict_from_json(file_path)

    @staticmethod
    def _print_retry_attempt(retry_state: RetryCallState):
        if retry_state.outcome is not None:
            last_exception = retry_state.outcome.exception()
            print(f"Próba nr {retry_state.attempt_number} nieudana. Czekam na ponowienie. Przyczyna: {last_exception}")

    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=2, min=61, max=366),
        before_sleep=_print_retry_attempt
    )
    def _gemini_request(
            self, prompt: str, attachments_paths: list[str], response_schema_path: str
    ) -> GenerateContentResponse:
        return self._get_client().models.generate_content(
            model=config.response_model,
            contents=self._get_contents(prompt, attachments_paths),
            config=self._get_config(self._get_response_schema(response_schema_path)),
        )


    def analyse_mail(self, prompt: str, mail_data: ResponseMailModel) -> Dict[str, Any]:

        try:

            print(f"Analizuję mail {mail_data.get_mail_title()} od {mail_data.get_mail_sender()}.")

            response = self._gemini_request(
                prompt, mail_data.get_attachments_paths(), config.mail_analysis_response_schema_path
            )

            print(f"Otrzymano odpowiedź od AI: {response.text}")

        except Exception as e:

            print(f"Wystąpił błąd podczas analizowania mail'a {mail_data.get_mail_title()} od "
                  f"{mail_data.get_mail_sender()}: {e}")
            return {"response_type": {"other_error": True, "additional_info": str(e)}}

        return self._type_converter.str_to_dict_or_list(response.text)

    def match_teryt(self, prompt: str, mail_data: ResponseMailModel) -> List[str]:

        try:

            print(f"Szukam kodu TERYT dla wiadomości {mail_data.get_mail_title()} od {mail_data.get_mail_sender()}.")

            response = self._gemini_request(
                prompt, mail_data.get_attachments_paths(), config.teryt_matcher_response_schema_path
            )

            print(f"Dopasowano kod TERYT: {response.text}")

        except Exception as e:

            print(f"Wystąpił błąd podczas dopasowywania kodu TERYT dla wiadomości {mail_data.get_mail_title()} od "
                  f"{mail_data.get_mail_sender()}: {e}")
            return []

        return self._type_converter.str_to_dict_or_list(response.text)