import re
from datetime import datetime

from src.enums import GenderEnum


class Utils:
    def get_salutation_denominator(self, authority_teryt: str, governor_gender_code: str) -> str:

        if governor_gender_code not in [GenderEnum.MALE.value, GenderEnum.FEMALE.value]:
            raise ValueError(f"Nieprawidłowy kod płci dla włodarza gminy o TERYT {authority_teryt}")

        if governor_gender_code == GenderEnum.MALE.value:
            return "Szanowny Pan"
        elif governor_gender_code == GenderEnum.FEMALE.value:
            return "Szanowna Pani"

    def get_salutation_vocative(self, authority_teryt: str, governor_gender_code: str) -> str:

        if governor_gender_code not in [GenderEnum.MALE.value, GenderEnum.FEMALE.value]:
            raise ValueError(f"Nieprawidłowy kod płci dla włodarza gminy o TERYT {authority_teryt}")

        if governor_gender_code == GenderEnum.MALE.value:
            return "Szanowny Panie"
        elif governor_gender_code == GenderEnum.FEMALE.value:
            return "Szanowna Pani"

    def get_today_date(self) -> str:
        return datetime.today().strftime('%d.%m.%Y')

    def check_email_format(self, email_str: str) -> bool:
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

        return all(re.match(email_regex, email) for email in email_str.split(';'))