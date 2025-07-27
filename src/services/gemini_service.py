from typing import Dict, List, Any, Union

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
            vertexai=True,
            project=config.project_id,
            location=config.location_code,
        )

    def _get_contents(self, prompt: str, attachments_paths: List[str]):

        attachments = []

        for attachment_path in attachments_paths:

            with open(attachment_path, "rb") as f:

                document_bytes = f.read()
                document = types.Part.from_bytes(document_bytes)
                attachments.append(document)

        return [
            types.Content(
                role="user",
                parts=[
                    prompt,
                    *attachments
                ]
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

    def analyse_mail(self, prompt: str, mail_data: ResponseMailModel) -> Dict[str, Any]:

        print(prompt)
        print(mail_data.get_attachments_paths())

        response = self._get_client().models.generate_content(
            model=config.response_model,
            contents=self._get_contents(prompt, mail_data.get_attachments_paths()),
            config=self._get_config(config.mail_analysis_response_schema_path),
        )

        return self._type_converter.str_to_dict_or_list(response.text)

    def match_teryt(self, prompt: str, mail_data: ResponseMailModel) -> List[str]:

        response = self._get_client().models.generate_content(
            model=config.response_model,
            contents=self._get_contents(prompt, mail_data.get_attachments_paths()),
            config=self._get_config(config.teryt_matcher_response_schema_path),
        )

        return self._type_converter.str_to_dict_or_list(response.text)