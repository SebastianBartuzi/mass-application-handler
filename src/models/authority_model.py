from src.enums import GenderEnum
from src.utils import PathCreator, Utils

class AuthorityModel:

    _authority_teryt: str = ""
    _authority_name: str = ""
    _governor_title: str = ""
    _governor_gender: str = ""
    _governor_name: str = ""
    _authority_office_name: str = ""
    _authority_office_mail: str = ""
    _application_docx_path: str = ""
    _application_pdf_path: str = ""

    def __init__(self) -> None:
        self._utils = Utils()
        self._path_creator = PathCreator()

    def set_authority_teryt(self, authority_teryt: int) -> None:
        if not isinstance(authority_teryt, int) or not authority_teryt:
            raise ValueError(f"Nieprawidłowa wartość TERYT: {authority_teryt}!")

        authority_teryt_str = str(authority_teryt)
        if len(authority_teryt_str) == 5:
            authority_teryt_str = f"0{authority_teryt_str}"

        self._authority_teryt = authority_teryt_str
        self._application_docx_path = self._path_creator.get_application_docx_path(self._authority_teryt)
        self._application_pdf_path = self._path_creator.get_application_pdf_path(self._authority_teryt)

    def get_authority_teryt(self) -> str:
        return self._authority_teryt

    def set_authority_name(self, authority_name: str) -> None:
        if not isinstance(authority_name, str) or not authority_name:
            raise ValueError(f"Nieprawidłowa wartość authority_name {authority_name} dla TERYT "
                             f"{self._authority_teryt}!")

        self._authority_name = authority_name

    def get_authority_name(self) -> str:
        return self._authority_name

    def set_governor_title(self, governor_title: str) -> None:
        if not isinstance(governor_title, str) or not governor_title:
            raise ValueError(f"Nieprawidłowa wartość governor_title {governor_title} dla TERYT "
                             f"{self._authority_teryt}!")

        self._governor_title = governor_title

    def get_governor_title(self) -> str:
        return self._governor_title

    def set_governor_gender(self, governor_gender: str) -> None:
        if governor_gender not in [GenderEnum.MALE.value, GenderEnum.FEMALE.value]:
            raise ValueError(f"Nieprawidłowa wartość governor_gender {governor_gender} dla TERYT "
                             f"{self._authority_teryt}!")

        self._governor_gender = governor_gender

    def get_governor_gender(self) -> str:
        return self._governor_gender

    def set_governor_name(self, governor_name: str) -> None:
        if not isinstance(governor_name, str) or not governor_name:
            raise ValueError(f"Nieprawidłowa wartość governor_name {governor_name} dla TERYT {self._authority_teryt}!")

        self._governor_name = governor_name

    def get_governor_name(self) -> str:
        return self._governor_name

    def set_authority_office_name(self, authority_office_name: str) -> None:
        if not isinstance(authority_office_name, str) or not authority_office_name:
            raise ValueError(f"Nieprawidłowa wartość authority_office_name {authority_office_name} dla TERYT "
                             f"{self._authority_teryt}!")

        self._authority_office_name = authority_office_name

    def get_authority_office_name(self) -> str:
        return self._authority_office_name

    def set_authority_office_mail(self, authority_office_mail: str) -> None:
        if not isinstance(authority_office_mail, str) or not authority_office_mail:
            raise ValueError(f"Nieprawidłowa wartość authority_office_mail {authority_office_mail} dla TERYT "
                             f"{self._authority_teryt}!")

        authority_office_mail = (authority_office_mail.strip().replace(" ", "")
                                 .replace(",", ";")
                                 .replace("\n", ";"))

        if not self._utils.check_email_format(authority_office_mail):
            raise ValueError(f"Nieprawidłowy format authority_office_mail {authority_office_mail} dla TERYT "
                             f"{self._authority_teryt}!")

        if ";" in authority_office_mail:
            print(f"Więcej niż jeden adres e-mail dla TERYT {self._authority_teryt}.")

        self._authority_office_mail = authority_office_mail

    def get_authority_office_mail(self) -> str:
        return self._authority_office_mail

    def get_application_docx_path(self) -> str:
        return self._application_docx_path

    def get_application_pdf_path(self) -> str:
        return self._application_pdf_path